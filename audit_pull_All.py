import msal
import requests
import os
import re
import sys
import json
import warnings
import time
from datetime import datetime
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

def is_headless():
    """
    Detect if running in a headless environment with no browser available.
    Checks for SSH sessions and Linux systems without a display server.
    """
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
        # Device flow — prints a short code the user enters at microsoft.com/devicelogin
        # Works on any headless server; authenticate from any device with a browser
        print("Headless environment detected — using device flow authentication.")
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            print(f"Error initiating device flow: {flow}")
            return None
        print(f"\n{flow['message']}\n")   # prints the code and URL
        result = app.acquire_token_by_device_flow(flow)
    else:
        print("Opening browser for interactive login...")
        result = app.acquire_token_interactive(scopes=scopes)

    if "access_token" not in result:
        print(f"Auth failed: {result.get('error_description')}")
        return None
    return result["access_token"]


# ---------------------------------------------------------------------------
# GRAPH HELPERS
# ---------------------------------------------------------------------------

CALL_DELAY        = 2.0   # seconds between every API call — ~30 req/min
REQUEST_LIMIT     = 80    # proactive pause after this many requests
REQUEST_PAUSE     = 120   # seconds to pause when limit is reached
_request_counter  = 0     # tracks total API calls made this session
_consecutive_429s = 0     # tracks back-to-back 429s across all calls


def safe_log(msg):
    """Log to screen and file if the log file is open, otherwise just print."""
    print(msg)
    try:
        _log_file.write(msg + "\n")
    except Exception:
        pass  # log file may not be open yet during early auth calls


def graph_get(url, headers, retries=6):
    """GET a Graph URL with a courtesy delay + exponential backoff on errors."""
    global _request_counter, _consecutive_429s
    _request_counter += 1

    # Proactive pause every REQUEST_LIMIT calls — prevents hitting the rate limit
    # inside large notebooks before Microsoft has a chance to throttle us
    if _request_counter % REQUEST_LIMIT == 0:
        print(f"  [Proactive pause — {_request_counter} requests made. Waiting {REQUEST_PAUSE}s...]")
        time.sleep(REQUEST_PAUSE)

    time.sleep(CALL_DELAY)  # courtesy delay on every call
    for attempt in range(retries):
        try:
            resp = requests.get(url, headers=headers, timeout=30)

            if resp.status_code == 200:
                _consecutive_429s = 0   # reset on any success
                return resp

            elif resp.status_code == 429:
                _consecutive_429s += 1
                body_preview = resp.text[:500]
                retry_after = resp.headers.get("Retry-After")
                wait = int(retry_after) if retry_after else 30 * (2 ** attempt)
                print(f"  [429 Rate limited — attempt {attempt+1}/{retries}] Waiting {wait}s")
                print(f"  Server message: {body_preview}")

                # Two consecutive 429s means the backoff isn't clearing the throttle
                if _consecutive_429s >= 2:
                    safe_log(f"\n[FATAL] Received {_consecutive_429s} consecutive 429 errors.")
                    safe_log(f"[FATAL] Microsoft is still throttling after waiting. Exiting gracefully.")
                    safe_log(f"[FATAL] Resume by re-running the script — already-archived pages will be skipped.")
                    try:
                        _log_file.close()
                    except Exception:
                        pass
                    sys.exit(1)

                time.sleep(wait)

            elif resp.status_code == 401:
                _consecutive_429s = 0
                print(f"  [401 Unauthorized] Token expired or missing permission.")
                print(f"  URL: {url}")
                print(f"  Response: {resp.text[:300]}")
                return None

            else:
                _consecutive_429s = 0
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
    Download all images, drawings, and attachments from a OneNote page.
    - Images: <img src> and <img data-fullres-src> (prefers full-res)
    - Attachments: <object data-attachment> (Word, PowerPoint, PDF, etc.)
    Updates the soup in-place so saved HTML points to local files.
    """
    os.makedirs(media_dir, exist_ok=True)
    attachments_dir = os.path.join(os.path.dirname(media_dir), "attachments")
    media_count = 0
    attachment_count = 0

    # --- Images ---
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

    # --- Attachments (Word, PowerPoint, PDF, etc.) ---
    for obj in soup.find_all('object', attrs={'data-attachment': True}):
        filename = re.sub(r'[\\/*?:"<>|]', "", obj.get('data-attachment', 'attachment'))
        attach_url = obj.get('data', '')

        if not attach_url.startswith('http'):
            continue

        resp = graph_get(attach_url, headers)
        if resp is None:
            print(f"    [Attachment skip] Could not download: {filename}")
            continue

        os.makedirs(attachments_dir, exist_ok=True)
        filepath = os.path.join(attachments_dir, filename)

        with open(filepath, "wb") as f:
            f.write(resp.content)

        # Update the object tag to point locally
        obj['data'] = f"attachments/{filename}"
        attachment_count += 1
        print(f"    [Attachment] Saved: {filename}")

    return media_count, attachment_count


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

token = get_access_token()
if not token:
    exit(1)

headers = {'Authorization': f'Bearer {token}'}
root_audit_dir = os.path.join("onenote_audit", "01_Raw_Audit")
os.makedirs(root_audit_dir, exist_ok=True)

# --- Log file setup ---
# Named with timestamp so each run gets its own file — never overwritten
log_path = os.path.join(
    "onenote_audit",
    f"archive_log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"
)
_log_file = open(log_path, "w", encoding="utf-8", buffering=1)  # buffering=1 = line-buffered (writes instantly)
print(f"Logging to: {log_path}\n")

def log(msg=""):
    """Print to screen and write to log file immediately."""
    print(msg)
    _log_file.write(msg + "\n")

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

# --- Notebook selection ---
print("Available notebooks:")
for i, nb in enumerate(notebooks, 1):
    print(f"  {i:>2}. {nb['displayName']}")

print()
print("  s <numbers> — skip these notebooks       e.g. s 1 3 5")
print("  p <numbers> — pull only these notebooks  e.g. p 2 4")
print("  Enter       — process all notebooks")
print()

selection = input("Your choice: ").strip().lower()

if selection.startswith("s "):
    skip_nums = set(int(x) for x in selection[2:].split() if x.isdigit())
    notebooks = [nb for i, nb in enumerate(notebooks, 1) if i not in skip_nums]
    print(f"\nSkipping {len(skip_nums)} notebook(s). Processing {len(notebooks)}.\n")
elif selection.startswith("p "):
    pull_nums = set(int(x) for x in selection[2:].split() if x.isdigit())
    notebooks = [nb for i, nb in enumerate(notebooks, 1) if i in pull_nums]
    print(f"\nPulling {len(notebooks)} notebook(s).\n")
else:
    print(f"\nProcessing all {len(notebooks)} notebooks.\n")

# Log the final selection
log(f"Notebooks selected for this run:")
for nb in notebooks:
    log(f"  - {nb['displayName']}")
log()

total_pages       = 0
total_images      = 0
total_attachments = 0
total_errors      = 0

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

    # Per-notebook counters
    nb_pages       = 0
    nb_images      = 0
    nb_attachments = 0
    nb_errors      = 0

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
                nb_pages += 1
                continue

            # --- Fetch HTML content ---
            content_resp = graph_get(page['contentUrl'], headers)
            if content_resp is None:
                print(f"    [FAILED] {nb_name} > {section_name} > {title}")
                nb_errors += 1
                continue

            soup = BeautifulSoup(content_resp.text, 'html.parser')

            # --- Download all media and attachments ---
            media_count, attachment_count = download_media(soup, media_dir, headers)

            # --- Save updated HTML ---
            with open(os.path.join(page_dir, "index.html"), "w", encoding='utf-8') as f:
                f.write(str(soup))

            nb_pages       += 1
            nb_images      += media_count
            nb_attachments += attachment_count
            print(f"    Archived [{media_count} media, {attachment_count} attachments]: {section_name} > {title}")

    # --- Per-notebook summary — written to screen and log immediately ---
    log(f"\n  ┌─ Summary: {nb_name}")
    log(f"  │  Sections     : {len(sections)}")
    log(f"  │  Pages        : {nb_pages}")
    log(f"  │  Images       : {nb_images}")
    log(f"  │  Attachments  : {nb_attachments}")
    log(f"  └─ Errors       : {nb_errors}\n")

    # Accumulate into grand totals
    total_pages       += nb_pages
    total_images      += nb_images
    total_attachments += nb_attachments
    total_errors      += nb_errors

# --- Grand total summary ---
log(f"\n{'=' * 40}")
log(f"  ARCHIVE COMPLETE")
log(f"{'=' * 40}")
log(f"  Notebooks   : {len(notebooks)}")
log(f"  Pages       : {total_pages}")
log(f"  Images      : {total_images}")
log(f"  Attachments : {total_attachments}")
log(f"  Errors      : {total_errors}")
log(f"  Output      : {os.path.abspath(root_audit_dir)}")
_log_file.close()

