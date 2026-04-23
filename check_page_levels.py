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

# Fetch ALL pages with $select — the only way level/order are returned
all_pages = []
url = f"{BASE}/sections/{cand_sec['id']}/pages?$select=id,title,level,order&$top=100"
while url:
    resp = requests.get(url, headers=headers).json()
    all_pages.extend(resp.get("value", []))
    url  = resp.get("@odata.nextLink")

print(f"\n{len(all_pages)} pages in Corelight > Candidates (with $select=level,order):")
print(f"{'level':<8} {'order':<20} title")
print("-" * 60)
for p in all_pages:
    level = p.get("level")
    order = p.get("order")
    marker = " ◄" if p["title"].lower() in ("yes","no","maybe","in-process","in process","pipeline") else ""
    print(f"{str(level) if level is not None else '(absent)':<8} {str(order) if order is not None else '(absent)':<20} {p['title']}{marker}")
