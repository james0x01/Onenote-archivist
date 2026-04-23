#!/usr/bin/env python3
"""
patch_manifest_hierarchy.py — Backfill 'level' and 'order' into existing manifests.

Why this exists
---------------
audit_pull_All.py only updates a manifest entry when it downloads page content
(i.e. the page is new or changed).  Pages that haven't changed since the last
pull keep their old manifest entry, which pre-dates the level/order fields.

This script hits the Graph API for page *metadata only* (no HTML download),
then writes level and order into every manifest entry — including unchanged pages.
It is safe to run at any time; it never touches 01_Raw_Audit content files.

Speed: ~2s per page (API rate-limit courtesy delay) — same metadata calls that
audit_pull_All.py makes, but without any content downloads.

Usage
-----
    python3 patch_manifest_hierarchy.py          # menu to pick notebooks
    python3 patch_manifest_hierarchy.py --all    # patch every notebook
"""

import os
import re
import sys
import json
import time
import argparse
import warnings
from datetime import datetime
from pathlib import Path

import msal
import requests
from dotenv import load_dotenv

warnings.filterwarnings("ignore", category=UserWarning, module='urllib3')

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

load_dotenv()
CLIENT_ID = os.getenv("ONENOTE_CLIENT_ID")
if not CLIENT_ID:
    print("Error: ONENOTE_CLIENT_ID not found in .env")
    sys.exit(1)

GRAPH_BASE    = "https://graph.microsoft.com/v1.0"
CALL_DELAY    = 1.5   # seconds between calls — lighter than full pull
RAW_DIR       = Path("onenote_audit") / "01_Raw_Audit"

# ---------------------------------------------------------------------------
# AUTH  (same device-flow logic as audit_pull_All.py)
# ---------------------------------------------------------------------------

def is_headless():
    if os.environ.get('SSH_CLIENT') or os.environ.get('SSH_TTY'):
        return True
    if sys.platform.startswith('linux'):
        if not os.environ.get('DISPLAY') and not os.environ.get('WAYLAND_DISPLAY'):
            return True
    return False


def get_access_token():
    authority = "https://login.microsoftonline.com/consumers"
    scopes    = ["User.Read", "Notes.Read", "Notes.Read.All"]
    app       = msal.PublicClientApplication(CLIENT_ID, authority=authority)
    for account in app.get_accounts():
        app.remove_account(account)

    if is_headless():
        print("Headless environment — using device flow.")
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            print(f"Device flow error: {flow}")
            return None
        print(f"\n{flow['message']}\n")
        result = app.acquire_token_by_device_flow(flow)
    else:
        print("Opening browser for interactive login...")
        result = app.acquire_token_interactive(scopes=scopes)

    if "access_token" not in result:
        print(f"Auth failed: {result.get('error_description')}")
        return None
    return result["access_token"]


# ---------------------------------------------------------------------------
# GRAPH HELPERS  (lightweight — no backoff complexity needed for metadata)
# ---------------------------------------------------------------------------

_request_count = 0

def graph_get(url, headers):
    """GET with courtesy delay and simple 429/5xx retry."""
    global _request_count
    _request_count += 1
    time.sleep(CALL_DELAY)

    for attempt in range(5):
        try:
            resp = requests.get(url, headers=headers, timeout=30)
            if resp.status_code == 200:
                return resp
            elif resp.status_code == 429:
                wait = int(resp.headers.get("Retry-After", 30 * (2 ** attempt)))
                print(f"  [429] Waiting {wait}s ...")
                time.sleep(wait)
            elif resp.status_code in (500, 502, 503, 504):
                wait = 15 * (2 ** attempt)
                print(f"  [HTTP {resp.status_code}] Retrying in {wait}s ...")
                time.sleep(wait)
            else:
                print(f"  [HTTP {resp.status_code}] {url[:80]}")
                return None
        except requests.exceptions.RequestException as e:
            wait = 10 * (2 ** attempt)
            print(f"  [Network error] {e} — retrying in {wait}s")
            time.sleep(wait)
    return None


def get_all_items(url, headers):
    """Paginated Graph fetch — follows @odata.nextLink."""
    items    = []
    next_url = url + ("&" if "?" in url else "?") + "$top=100"
    while next_url:
        resp = graph_get(next_url, headers)
        if resp is None:
            break
        data     = resp.json()
        items.extend(data.get("value", []))
        next_url = data.get("@odata.nextLink")
    return items


def get_sections_recursive(notebook_id, headers):
    """Flat list of all sections (including nested section groups), each with '_path'."""
    def fetch(group_id, group_path, is_notebook=False):
        prefix  = "notebooks" if is_notebook else "sectionGroups"
        base_id = notebook_id if is_notebook else group_id
        sec_url = (f"{GRAPH_BASE}/me/onenote/notebooks/{base_id}/sections"
                   if is_notebook else
                   f"{GRAPH_BASE}/me/onenote/sectionGroups/{group_id}/sections")
        grp_url = (f"{GRAPH_BASE}/me/onenote/notebooks/{base_id}/sectionGroups"
                   if is_notebook else
                   f"{GRAPH_BASE}/me/onenote/sectionGroups/{group_id}/sectionGroups")
        result = []
        for s in get_all_items(sec_url, headers):
            s["_path"] = group_path
            result.append(s)
        for g in get_all_items(grp_url, headers):
            gname = re.sub(r'[\\/*?:"<>|]', "", g["displayName"])
            result.extend(fetch(g["id"], os.path.join(group_path, gname)))
        return result

    return fetch(notebook_id, "", is_notebook=True)


# ---------------------------------------------------------------------------
# MANIFEST HELPERS
# ---------------------------------------------------------------------------

def load_manifest(nb_dir: Path) -> dict:
    path = nb_dir / "manifest.json"
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"pages": {}}


def save_manifest(nb_dir: Path, manifest: dict):
    path = nb_dir / "manifest.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

parser = argparse.ArgumentParser(description="Backfill level/order into manifests")
parser.add_argument("--all", action="store_true", help="Patch all notebooks without prompting")
args = parser.parse_args()

print("=" * 60)
print("  patch_manifest_hierarchy — backfill level & order")
print("=" * 60)
print()

token = get_access_token()
if not token:
    sys.exit(1)
headers = {"Authorization": f"Bearer {token}"}

# Verify token
me = graph_get(f"{GRAPH_BASE}/me", headers)
if me is None:
    print("Token check failed. Exiting.")
    sys.exit(1)
print(f"Signed in as: {me.json().get('displayName')}\n")

# Fetch notebooks
print("Fetching notebooks ...")
notebooks = get_all_items(f"{GRAPH_BASE}/me/onenote/notebooks", headers)
print(f"Found {len(notebooks)} notebook(s).\n")

if args.all:
    selected = notebooks
else:
    for i, nb in enumerate(notebooks, 1):
        print(f"  {i:>2}. {nb['displayName']}")
    print()
    print("  a           — patch all")
    print("  p <numbers> — patch only these  e.g. p 1 3")
    print()
    choice = input("Your choice: ").strip().lower()
    if choice == "a" or choice == "":
        selected = notebooks
    elif choice.startswith("p "):
        nums     = {int(x) for x in choice[2:].split() if x.isdigit()}
        selected = [nb for i, nb in enumerate(notebooks, 1) if i in nums]
    else:
        print("Unrecognised — patching all.")
        selected = notebooks

print(f"\nPatching {len(selected)} notebook(s):\n")

total_patched  = 0
total_skipped  = 0
total_added    = 0

for nb in selected:
    nb_name  = re.sub(r'[\\/*?:"<>|]', "", nb["displayName"])
    nb_dir   = RAW_DIR / nb_name
    manifest = load_manifest(nb_dir)

    print(f"  Notebook: {nb_name}")

    sections = get_sections_recursive(nb["id"], headers)
    print(f"    {len(sections)} section(s)")

    if not sections:
        print(f"    [Skip — no sections found (shared or empty notebook)]\n")
        continue

    nb_patched = 0
    nb_added   = 0

    for section in sections:
        sec_name  = re.sub(r'[\\/*?:"<>|]', "", section["displayName"])
        sec_path  = section["_path"]

        # Fetch page metadata — level, order, title, lastModified
        pages = get_all_items(
            f"{GRAPH_BASE}/me/onenote/sections/{section['id']}/pages"
            f"?$select=id,title,lastModifiedDateTime,level,order",
            headers
        )

        for idx, page in enumerate(pages):
            title     = re.sub(r'[\\/*?:"<>|]', "", page.get("title") or "Untitled_Page")
            page_key  = "/".join(filter(None, [
                sec_path.replace("\\", "/"), sec_name, title
            ]))
            # level/order are absent for personal Microsoft accounts.
            # Fall back: level=0 (top-level), order=response index (visual sequence).
            level     = page.get("level") if page.get("level") is not None else 0
            order     = page.get("order") if page.get("order") is not None else idx

            existing = manifest.get("pages", {}).get(page_key)

            if existing is None:
                # Page not yet in manifest — add a stub (no content pulled)
                manifest.setdefault("pages", {})[page_key] = {
                    "lastModifiedDateTime": page.get("lastModifiedDateTime", ""),
                    "pulled":  None,   # content not yet downloaded
                    "level":   level,
                    "order":   order,
                }
                nb_added += 1
            else:
                # Update level/order in the existing entry
                existing["level"] = level
                existing["order"] = order
                nb_patched += 1

    save_manifest(nb_dir, manifest)
    total_patched += nb_patched
    total_added   += nb_added
    print(f"    ✓ {nb_patched} entries updated, {nb_added} stubs added\n")

print("=" * 60)
print(f"  PATCH COMPLETE")
print(f"  Entries updated : {total_patched}")
print(f"  Stubs added     : {total_added}")
print(f"  API calls made  : {_request_count}")
print("=" * 60)
