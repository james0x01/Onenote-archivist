#!/usr/bin/env python3
"""
check_page_levels.py — Diagnostic: show raw level/order from OneNote API.

Usage:
    python3 check_page_levels.py
"""
import requests
import msal
import os
from dotenv import dotenv_values

cfg       = dotenv_values(os.path.join(os.path.dirname(__file__), ".env"))
CLIENT_ID = cfg.get("ONENOTE_CLIENT_ID")

app  = msal.PublicClientApplication(CLIENT_ID, authority="https://login.microsoftonline.com/consumers")
flow = app.initiate_device_flow(scopes=["Notes.Read", "Notes.Read.All"])
print(flow["message"])
result  = app.acquire_token_by_device_flow(flow)
headers = {"Authorization": f"Bearer {result['access_token']}"}

BASE = "https://graph.microsoft.com/v1.0/me/onenote"

# Find Corelight notebook
nbs       = requests.get(f"{BASE}/notebooks", headers=headers).json()
corelight = next(n for n in nbs["value"] if n["displayName"] == "Corelight")

# Find Candidates section
sections = requests.get(
    f"{BASE}/notebooks/{corelight['id']}/sections", headers=headers
).json()
cand_sec = next(s for s in sections["value"] if s["displayName"] == "Candidates")

# Fetch first 15 pages WITHOUT $select — see every field the API returns
pages = requests.get(
    f"{BASE}/sections/{cand_sec['id']}/pages?$top=15",
    headers=headers
).json()

print(f"\nRaw fields for first 15 pages in Corelight > Candidates:")
print(f"{'level':<8} {'order':<20} title")
print("-" * 60)
for p in pages["value"]:
    print(f"{str(p.get('level','?')):<8} {str(p.get('order','?')):<20} {p['title']}")
