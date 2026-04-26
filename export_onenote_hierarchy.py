#!/usr/bin/env python3
"""
export_onenote_hierarchy.py — Extract exact page hierarchy from OneNote desktop.

Uses the OneNote COM API (Windows only) to read the true level/order of every
page in every notebook, then writes rollup_groups.json files with 100% accurate
topic groupings.

Requirements:
    Windows + OneNote desktop app + notebooks synced
    pip install pywin32

Usage (on Windows):
    python export_onenote_hierarchy.py                # writes to .\rollup_groups\
    python export_onenote_hierarchy.py --out C:\path  # custom output directory

After running:
    Copy each rollup_groups.json to the Mac and place in Dropbox at:
        01_Raw_Audit/{Notebook}/rollup_groups.json
    Then SCP to Cerebro and run: python3 summarize_rollups.py --force
"""

import os
import re
import sys
import json
import argparse
from pathlib import Path
import xml.etree.ElementTree as ET

try:
    import win32com.client
except ImportError:
    print("ERROR: pywin32 not installed.  Run: pip install pywin32")
    sys.exit(1)

NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"

SKIP_SECTIONS = {"candidates"}

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def sanitize(name):
    """Remove characters that are illegal in file/folder names."""
    return re.sub(r'[\\/*?:"<>|]', "", name or "").strip()


def parse_hierarchy(xml_str):
    """
    Parse OneNote hierarchy XML and return:
        {notebook_name: {section_name: [{name, level, order}, ...]}}
    """
    root      = ET.fromstring(xml_str)
    ns        = {"one": NS}
    notebooks = {}

    for nb_el in root.findall(".//one:Notebook", ns):
        nb_name = sanitize(nb_el.get("name", "Unknown"))
        notebooks[nb_name] = {}

        # Walk all sections (including those inside section groups)
        for sec_el in nb_el.findall(".//one:Section", ns):
            sec_name = sanitize(sec_el.get("name", "Unknown"))

            if sec_name.lower() in SKIP_SECTIONS:
                continue

            pages = []
            for idx, pg_el in enumerate(sec_el.findall("one:Page", ns)):
                title = sanitize(pg_el.get("name", "Untitled_Page"))
                level = int(pg_el.get("pageLevel", 1)) - 1  # COM uses 1-based; convert to 0-based
                pages.append({
                    "name":  title,
                    "level": level,
                    "order": idx,
                })

            if pages:
                notebooks[nb_name][sec_name] = pages

    return notebooks


def build_groups(pages):
    """
    Walk level-0 / level-1 page list and return:
        {parent_name: [child_name, ...]}
    Level-0 pages with no level-1 children are placed in "Other".
    """
    groups   = {}
    other    = []
    current  = None

    for p in pages:
        if p["level"] == 0:
            current = p["name"]
            groups[current] = []
        else:
            if current:
                groups[current].append(p["name"])
            else:
                other.append(p["name"])

    # Fold childless level-0 pages into "Other"
    childless = [name for name, children in groups.items() if not children]
    for name in childless:
        other.append(name)
        del groups[name]

    if other:
        groups["Other"] = other

    return groups


# ---------------------------------------------------------------------------
# ARGUMENT PARSING
# ---------------------------------------------------------------------------

parser = argparse.ArgumentParser(
    description="Export OneNote page hierarchy to rollup_groups.json files"
)
parser.add_argument(
    "--out", default="rollup_groups",
    help="Output directory (default: .\\rollup_groups)"
)
parser.add_argument(
    "--notebook", default=None,
    help="Only export this notebook (default: all)"
)
args = parser.parse_args()

out_dir = Path(args.out)
out_dir.mkdir(parents=True, exist_ok=True)

print("=" * 60)
print("  export_onenote_hierarchy — OneNote COM API")
print("=" * 60)
print()
print("Connecting to OneNote...")

try:
    onenote = win32com.client.Dispatch("OneNote.Application")
except Exception as e:
    print(f"ERROR: Could not connect to OneNote: {e}")
    print("Make sure OneNote desktop is installed and has run at least once.")
    sys.exit(1)

# HierarchyScope: hsPages = 4  (includes notebooks → sections → pages)
print("Fetching full hierarchy (this may take a moment)...")
try:
    xml_str = onenote.GetHierarchy("", 4, "")
except Exception as e:
    print(f"ERROR: GetHierarchy failed: {e}")
    sys.exit(1)

print("Parsing hierarchy...\n")
notebooks = parse_hierarchy(xml_str)

if args.notebook:
    notebooks = {k: v for k, v in notebooks.items() if k == args.notebook}
    if not notebooks:
        print(f"ERROR: Notebook '{args.notebook}' not found.")
        print(f"Available: {', '.join(parse_hierarchy(xml_str).keys())}")
        sys.exit(1)

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

total_notebooks = 0
total_groups    = 0

for nb_name, sections in sorted(notebooks.items()):
    print(f"Notebook: {nb_name}")
    config = {}

    for sec_name, pages in sorted(sections.items()):
        groups = build_groups(pages)
        if not groups:
            continue

        config[sec_name] = groups
        n_groups  = len([g for g in groups if g != "Other"])
        n_other   = len(groups.get("Other", []))
        print(f"  {sec_name}: {n_groups} group(s), {n_other} ungrouped")
        total_groups += n_groups

    if not config:
        print("  (no sections with hierarchy)")
        continue

    # Write rollup_groups.json for this notebook
    nb_out = out_dir / nb_name
    nb_out.mkdir(parents=True, exist_ok=True)
    out_file = nb_out / "rollup_groups.json"
    out_file.write_text(json.dumps(config, indent=2), encoding="utf-8")
    print(f"  → Written: {out_file}\n")
    total_notebooks += 1

print("=" * 60)
print(f"  EXPORT COMPLETE")
print(f"  Notebooks : {total_notebooks}")
print(f"  Groups    : {total_groups}")
print(f"  Output    : {out_dir.resolve()}")
print("=" * 60)
print()
print("Next steps:")
print("  1. Copy each rollup_groups.json to Mac Dropbox:")
print("       01_Raw_Audit/{Notebook}/rollup_groups.json")
print("  2. SCP to Cerebro:")
print("       scp <file> lab-user@10.254.254.48:Onenote-archivist/onenote_audit/01_Raw_Audit/{Notebook}/")
print("  3. On Cerebro: python3 summarize_rollups.py --force")
