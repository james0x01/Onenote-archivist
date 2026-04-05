import msal
import requests
import os
import re
import json
import warnings
import webbrowser
from bs4 import BeautifulSoup

# Suppress SSL/LibreSSL warnings
warnings.filterwarnings("ignore", category=UserWarning, module='urllib3')

# --- CONFIGURATION ---
CLIENT_ID = 'INSERT CLIENT ID'
NOTEBOOK_NAME = "Book-Idea"

def get_access_token():
    # 'consumers' for personal accounts
    authority = "https://login.microsoftonline.com/consumers"
    scopes = ["User.Read", "Notes.Read", "Notes.Read.All"]
    
    app = msal.PublicClientApplication(CLIENT_ID, authority=authority)

    # Wipe cache for a fresh login
    for account in app.get_accounts():
        app.remove_account(account)

    print("Opening browser for one-click login...")
    
    # --- CHANGE: INTERACTIVE FLOW ---
    # This will open the browser and listen for the success message automatically.
    # No code entry required.
    result = app.acquire_token_interactive(scopes=scopes)

    if "access_token" in result:
        print("\nToken acquired successfully!")
        return result["access_token"]
    else:
        print(f"Auth Error: {result.get('error_description')}")
        return None

# --- EXECUTION ---
token = get_access_token()

if not token:
    print("Authentication failed. Exiting.")
    exit(1)

headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}

# Folder structure: onenote_audit/01_Raw_Audit/Book-Idea
base_dir = os.path.join("onenote_audit", "01_Raw_Audit", NOTEBOOK_NAME)
os.makedirs(base_dir, exist_ok=True)

print(f"\nSearching for notebook: {NOTEBOOK_NAME}...")
api_url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks"
nb_resp = requests.get(api_url, headers=headers).json()

try:
    notebook = next(nb for nb in nb_resp['value'] if nb['displayName'] == NOTEBOOK_NAME)
    notebook_id = notebook['id']
    
    # Fetch sections to build the hierarchy
    sections_url = f"https://graph.microsoft.com/v1.0/me/onenote/notebooks/{notebook_id}/sections"
    sections = requests.get(sections_url, headers=headers).json().get('value', [])

    for section in sections:
        # Clean section name for folder safety
        section_name = re.sub(r'[\\/*?:"<>|]', "", section['displayName'])
        
        # Get pages for this section
        sec_pages_url = f"https://graph.microsoft.com/v1.0/me/onenote/sections/{section['id']}/pages"
        pages = requests.get(sec_pages_url, headers=headers).json().get('value', [])

        for page in pages:
            # Clean title
            title = re.sub(r'[\\/*?:"<>|]', "", page.get('title') or "Untitled_Page")
            
            # Directory: Notebook/Section/Page
            page_dir = os.path.join(base_dir, section_name, title)
            os.makedirs(page_dir, exist_ok=True)

            # Download HTML content
            content_resp = requests.get(page['contentUrl'], headers=headers)
            if content_resp.status_code == 200:
                with open(os.path.join(page_dir, "index.html"), "w", encoding='utf-8') as f:
                    f.write(content_resp.text)
                print(f" Archived: {section_name} > {title}")

except StopIteration:
    print(f"Error: Notebook '{NOTEBOOK_NAME}' not found.")

print(f"\n--- PROCESS COMPLETE ---")
print(f"Audit located at: {os.path.abspath(base_dir)}")
