import msal
import requests
import os
import re
import json
import warnings
import time
from bs4 import BeautifulSoup

warnings.filterwarnings("ignore", category=UserWarning, module='urllib3')

# Load CLIENT_ID from .env file — never hardcode secrets in source code
from dotenv import load_dotenv
load_dotenv()
CLIENT_ID = os.getenv("ONENOTE_CLIENT_ID")
if not CLIENT_ID:
    print("Error: ONENOTE_CLIENT_ID not found. Create a .env file with that value.")
    exit(1)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


# ---------------------------------------------------------------------------
# AUTH
# ---------------------------------------------------------------------------

def get_access_token():
    authority = "https://login.microsoftonline.com/consumers"
    scopes = ["User.Read", "Notes.Read", "Notes.Read.All"]
    app = msal.PublicClientApplication(CLIENT_ID, authority=authority)
    for account in app.get_accounts():
        app.remove_account(account)
    print("Opening browser for interactive login...")
    result = app.acquire_token_interactive(scopes=scopes)
    if "access_token" not in result:
        print(f"Auth failed: {result.get('error_description')}")
        return None
    return result["access_token"]


# ---------------------------------------------------------------------------
# GRAPH HELPERS
# ---------------------------------------------------------------------------

CALL_DELAY = 1.5  # seconds between every API call — ~40 req/min, well under Graph limits

def graph_get(url, headers, retries=6):
    """GET a Graph URL with a courtesy delay + exponential backoff on errors."""
    time.sleep(CALL_DELAY)  # throttle every call, including media downloads
    for attempt in range(retries):
        try:
            resp = requests.get(url, headers=headers, timeout=30)

            if resp.status_code == 200:
                return resp

            elif resp.status_code == 429:
                # Always print the body — Microsoft often says exactly how long to wait
                body_preview = resp.text[:500]
                # Honour Retry-After if present; otherwise exponential backoff (30, 60, 120…)
                retry_after = resp.headers.get("Retry-After")
                wait = int(retry_after) if retry_after else 30 * (2 ** attempt)
                print(f"  [429 Rate limited — attempt {attempt+1}/{retries}] Waiting {wait}s")
                print(f"  Server message: {body_preview}")
                time.sleep(wait)

            elif resp.status_code == 401:
                print(f"  [401 Unauthorized] Token expired or missing permission.")
                print(f"  URL: {url}")
                print(f"  Response: {resp.text[:300]}")
                return None

            else:
                print(f"  [HTTP {resp.status_code}] {url}")
                print(f"  Response: {resp.text[:300]}")
                return None

        except requests.exceptions.RequestException as e:
            wait = 10 * (2 ** attempt)
            print(f"  [Network error attempt {attempt+1}/{retries}] {e} — retrying in {wait}s")
            time.sleep(wait)

    print(f"  [FAILED after {retries} attempts] {url}")
    return None


def get_all_pages(url, headers):
    """
    Fetch ALL items from a paginated Graph endpoint.
    Follows @odata.nextLink automatically.
    Appends $top=100 to minimise the number of pagination round-trips.
    """
    items = []
    # Add $top=100 only on the first call; nextLink already encodes its own params
    next_url = url + ("&" if "?" in url else "?") + "$top=100"
    page_num = 0
    while next_url:
        page_num += 1
        resp = graph_get(next_url, headers)
        if resp is None:
            print(f"  [Pagination stopped at page {page_num}]")
            break
        data = resp.json()
        batch = data.get('value', [])
        items.extend(batch)
        next_url = data.get('@odata.nextLink')  # None if no more pages
    return items


def get_sections_recursive(notebook_id, headers):
    """
    Fetch ALL sections in a notebook, including those inside section groups
    (section groups can be nested, so this recurses).
    Returns a flat list of section dicts, each with a '_path' key.
    """
    def fetch_sections_in_group(group_id, group_path, is_notebook=False):
        if is_notebook:
            sections_url = f"{GRAPH_BASE}/me/onenote/notebooks/{group_id}/sections"
            groups_url = f"{GRAPH_BASE}/me/onenote/notebooks/{group_id}/sectionGroups"
        else:
            sections_url = f"{GRAPH_BASE}/me/onenote/sectionGroups/{group_id}/sections"
            groups_url = f"{GRAPH_BASE}/me/onenote/sectionGroups/{group_id}/sectionGroups"

        found_sections = []
        for s in get_all_pages(sections_url, headers):
            s['_path'] = group_path
            found_sections.append(s)

        # Recurse into nested section groups
        for g in get_all_pages(groups_url, headers):
            group_name = re.sub(r'[\\/*?:"<>|]', "", g['displayName'])
            found_sections.extend(
                fetch_sections_in_group(g['id'], os.path.join(group_path, group_name))
            )
        return found_sections

    return fetch_sections_in_group(notebook_id, "", is_notebook=True)


# ---------------------------------------------------------------------------
# MEDIA DOWNLOAD
# ---------------------------------------------------------------------------

def download_media(soup, media_dir, headers):
    """
    Download all images and drawings from a OneNote page.
    OneNote uses both 'src' and 'data-fullres-src'; we prefer the full-res version.
    Updates the soup in-place so saved HTML points to local files.
    """
    os.makedirs(media_dir, exist_ok=True)
    media_count = 0

    for i, img in enumerate(soup.find_all('img')):
        # Prefer full-res source; fall back to standard src
        img_url = img.get('data-fullres-src') or img.get('src') or ""

        if not img_url.startswith('http'):
            continue  # skip data URIs or empty

        resp = graph_get(img_url, headers)
        if resp is None:
            print(f"    [Media skip] Could not download image {i}")
            continue

        # Determine extension from Content-Type header
        content_type = resp.headers.get('Content-Type', 'image/png')
        ext_map = {
            'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif',
            'image/svg+xml': 'svg', 'image/webp': 'webp'
        }
        ext = ext_map.get(content_type.split(';')[0].strip(), 'bin')
        filename = f"resource_{i}.{ext}"
        filepath = os.path.join(media_dir, filename)

        with open(filepath, "wb") as f:
            f.write(resp.content)

        # Update both src attributes to point locally
        img['src'] = f"media/{filename}"
        if img.get('data-fullres-src'):
            img['data-fullres-src'] = f"media/{filename}"

        media_count += 1

    return media_count


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

token = get_access_token()
if not token:
    exit(1)

headers = {'Authorization': f'Bearer {token}'}
root_audit_dir = os.path.join("onenote_audit", "01_Raw_Audit")
os.makedirs(root_audit_dir, exist_ok=True)

# --- Verify token ---
me = graph_get(f"{GRAPH_BASE}/me", headers)
if me is None:
    print("Token validation failed. Exiting.")
    exit(1)
print(f"Signed in as: {me.json().get('displayName')}\n")

# --- Fetch ALL notebooks (paginated) ---
print("Fetching all notebooks...")
notebooks = get_all_pages(f"{GRAPH_BASE}/me/onenote/notebooks", headers)
print(f"Found {len(notebooks)} notebook(s).\n")

total_pages = 0
total_errors = 0

for nb_idx, notebook in enumerate(notebooks, 1):
    nb_name = re.sub(r'[\\/*?:"<>|]', "", notebook['displayName'])
    print(f"[{nb_idx}/{len(notebooks)}] Notebook: {nb_name}")

    # --- Fetch ALL sections (including inside section groups) ---
    sections = get_sections_recursive(notebook['id'], headers)
    print(f"  {len(sections)} section(s) found (including section groups)")

    if not sections:
        print("  [No sections found — skipping]")
        continue

    # Pause between notebooks to let Microsoft's rate limit window reset
    if nb_idx > 1:
        print(f"  [Pausing 3 minutes before next notebook to avoid throttling...]")
        time.sleep(180)

    for section in sections:
        section_name = re.sub(r'[\\/*?:"<>|]', "", section['displayName'])
        # _path is the subfolder path if inside section groups, else empty string
        section_path = os.path.join(nb_name, section['_path'], section_name)

        # --- Fetch ALL pages in this section (paginated) ---
        pages = get_all_pages(
            f"{GRAPH_BASE}/me/onenote/sections/{section['id']}/pages",
            headers
        )

        for page in pages:
            title = re.sub(r'[\\/*?:"<>|]', "", page.get('title') or "Untitled_Page")
            page_dir = os.path.join(root_audit_dir, section_path, title)
            media_dir = os.path.join(page_dir, "media")
            os.makedirs(page_dir, exist_ok=True)

            # --- RESUME: skip pages already downloaded ---
            if os.path.exists(os.path.join(page_dir, "index.html")):
                print(f"    [Skipped — already archived]: {section_name} > {title}")
                total_pages += 1
                continue

            # --- Fetch HTML content ---
            content_resp = graph_get(page['contentUrl'], headers)
            if content_resp is None:
                print(f"    [FAILED] {nb_name} > {section_name} > {title}")
                total_errors += 1
                continue

            soup = BeautifulSoup(content_resp.text, 'html.parser')

            # --- Download all media ---
            media_count = download_media(soup, media_dir, headers)

            # --- Save updated HTML ---
            with open(os.path.join(page_dir, "index.html"), "w", encoding='utf-8') as f:
                f.write(str(soup))

            total_pages += 1
            print(f"    Archived [{media_count} media]: {section_name} > {title}")

print(f"\n--- ARCHIVE COMPLETE ---")
print(f"  Pages archived : {total_pages}")
print(f"  Errors         : {total_errors}")
print(f"  Output folder  : {os.path.abspath(root_audit_dir)}")

