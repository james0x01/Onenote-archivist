#!/usr/bin/env python3
"""
cluster_rollup_groups.py — LLM-based topic clustering for rollup_groups.json.

For each notebook section, sends page titles + summary snippets to Gemini
and asks it to suggest topic groupings.  Outputs rollup_groups.json per
notebook to 01_Raw_Audit/{Notebook}/.

Run this on your Mac (reads Dropbox paths, uses Gemini API).
Review the output, correct any mistakes, then SCP to Cerebro and run
summarize_rollups.py --force to regenerate rollups with the new groups.

Usage:
    python3 cluster_rollup_groups.py           # menu to pick notebooks
    python3 cluster_rollup_groups.py --all     # all notebooks
    python3 cluster_rollup_groups.py --force   # overwrite existing groups
"""

import os
import re
import sys
import json
import time
import argparse
from pathlib import Path
from dotenv import dotenv_values
from google import genai

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

_dir           = os.path.dirname(os.path.abspath(__file__))
cfg            = dotenv_values(os.path.join(_dir, ".env"))
GEMINI_API_KEY = cfg.get("GEMINI_API_KEY")

DROPBOX_BASE = Path("/Users/james/Dropbox/Business/OneNote")
RAW_DIR      = DROPBOX_BASE / "01_Raw_Audit"
SUM_DIR      = DROPBOX_BASE / "03_Summaries"

SKIP_SECTIONS  = {"candidates"}
CALL_DELAY     = 2      # seconds between API calls
SNIPPET_LEN    = 400    # chars of summary body to include per page
MIN_SECTION_SZ = 4      # skip sections with fewer pages than this

GEMINI_MODELS = {
    "g": ("gemini-2.5-flash", "gemini-2.5-flash  (fast, cloud)"),
    "f": ("gemini-2.5-pro",   "gemini-2.5-pro    (more capacity, cloud)"),
}

_llm_model_name = "gemini-2.5-flash"   # updated by model menu below

# ---------------------------------------------------------------------------
# PROMPT
# ---------------------------------------------------------------------------

CLUSTER_PROMPT = """\
You are analysing page titles from a OneNote section to group related pages \
into topic clusters.

Rules:
- Group pages that are chapters of the same book, course, or series.
- Group pages that cover the same sub-topic or theme.
- Standalone pages that don't belong to any group should be placed in a \
group named "Other".
- Give each group a short, descriptive name (the book title, course name, \
or topic).
- Every page title must appear in exactly one group.
- Return ONLY valid JSON — no explanation, no markdown fences.

Format:
{{
  "Group Name": ["Page Title 1", "Page Title 2"],
  "Another Group": ["Page Title 3", "Page Title 4"],
  "Other": ["Standalone Page"]
}}

Section: {section}
Notebook: {notebook}

Pages (title | first ~400 chars of existing summary):
{page_list}
"""

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def parse_frontmatter(text):
    if text.startswith("---"):
        end = text.find("---", 3)
        if end != -1:
            return text[end + 3:].strip()
    return text.strip()


def load_manifest(nb_dir):
    path = nb_dir / "manifest.json"
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {"pages": {}}


def get_summary_snippet(nb_name, sec_key, page_name):
    """Return up to SNIPPET_LEN chars of the page summary body, or ''."""
    sum_file = SUM_DIR / nb_name / sec_key / f"{page_name}.md"
    if not sum_file.exists():
        return ""
    try:
        body = parse_frontmatter(sum_file.read_text(encoding="utf-8"))
        # Strip markdown headers for cleaner context
        body = re.sub(r"^#+\s.*$", "", body, flags=re.MULTILINE).strip()
        return body[:SNIPPET_LEN].replace("\n", " ")
    except Exception:
        return ""


def call_gemini(client, prompt):
    time.sleep(CALL_DELAY)
    for attempt in range(3):
        try:
            return client.models.generate_content(
                model=_llm_model_name, contents=prompt
            ).text.strip()
        except Exception as e:
            wait = 20 * (attempt + 1)
            print(f"  [Retry {attempt+1}/3 in {wait}s]: {e}")
            time.sleep(wait)
    return None


def parse_llm_json(response):
    """Extract and parse a JSON object from LLM response."""
    if response is None:
        return None
    # Strip markdown fences if present
    text = re.sub(r"```(?:json)?\s*", "", response).strip().rstrip("`").strip()
    # Find first { ... } block
    start = text.find("{")
    end   = text.rfind("}")
    if start == -1 or end == -1:
        return None
    try:
        return json.loads(text[start:end + 1])
    except json.JSONDecodeError:
        return None


def build_section_pages(manifest, nb_name):
    """
    Returns {sec_key: [page_name, ...]} skipping SKIP_SECTIONS.
    Sections with fewer than MIN_SECTION_SZ pages are also skipped.
    """
    sections = {}
    for key in manifest.get("pages", {}):
        parts    = key.rsplit("/", 1)
        sec_key  = parts[0] if len(parts) == 2 else ""
        page_name = parts[-1]
        sec_leaf = sec_key.split("/")[-1].lower() if sec_key else ""
        if sec_leaf in SKIP_SECTIONS:
            continue
        sections.setdefault(sec_key, []).append(page_name)

    return {
        k: v for k, v in sections.items()
        if len(v) >= MIN_SECTION_SZ
    }

# ---------------------------------------------------------------------------
# ARGUMENT PARSING
# ---------------------------------------------------------------------------

parser = argparse.ArgumentParser(description="LLM-based topic clustering for rollups")
parser.add_argument("--all",   action="store_true", help="Process all notebooks")
parser.add_argument("--force", action="store_true",
                    help="Overwrite existing rollup_groups.json files")
args = parser.parse_args()

print("=" * 60)
print("  cluster_rollup_groups — LLM topic clustering")
print("=" * 60)
print()

if not GEMINI_API_KEY:
    print("ERROR: GEMINI_API_KEY not found in .env")
    sys.exit(1)

client = genai.Client(api_key=GEMINI_API_KEY)

# ---------------------------------------------------------------------------
# MODEL SELECTION
# ---------------------------------------------------------------------------

print("LLM model for clustering:")
for key, (_, label) in GEMINI_MODELS.items():
    print(f"  {key}  — {label}")
print()

model_choice = input("Model choice [g]: ").strip().lower() or "g"
if model_choice in GEMINI_MODELS:
    _llm_model_name, model_label = GEMINI_MODELS[model_choice]
else:
    _llm_model_name, model_label = GEMINI_MODELS["g"]
print(f"  Using: {model_label}\n")

# ---------------------------------------------------------------------------
# NOTEBOOK SELECTION
# ---------------------------------------------------------------------------

notebooks = sorted([d.name for d in RAW_DIR.iterdir()
                    if d.is_dir() and (d / "manifest.json").exists()])

if args.all:
    selected = notebooks
else:
    print("Available notebooks:")
    for i, nb in enumerate(notebooks, 1):
        print(f"  {i:>2}. {nb}")
    print()
    print("  a           — all")
    print("  p <numbers> — only these   e.g. p 1 3")
    print()
    choice = input("Your choice: ").strip().lower()
    if choice in ("a", ""):
        selected = notebooks
    elif choice.startswith("p "):
        nums     = {int(x) for x in choice[2:].split() if x.isdigit()}
        selected = [nb for i, nb in enumerate(notebooks, 1) if i in nums]
    else:
        print("Unrecognised — processing all.")
        selected = notebooks

print(f"\nProcessing {len(selected)} notebook(s): {', '.join(selected)}\n")

# ---------------------------------------------------------------------------
# MAIN LOOP
# ---------------------------------------------------------------------------

total_generated = 0
total_skipped   = 0
total_errors    = 0

for nb_name in selected:
    nb_dir   = RAW_DIR / nb_name
    out_path = nb_dir / "rollup_groups.json"

    if out_path.exists() and not args.force:
        # Check if it already has real groups (not just scaffolds)
        try:
            existing = json.loads(out_path.read_text(encoding="utf-8"))
            has_groups = any(
                isinstance(v, dict) and
                any(not k.startswith("_") for k in v)
                for v in existing.values()
            )
            if has_groups:
                print(f"[Skip — already has groups]: {nb_name}  (use --force to redo)")
                total_skipped += 1
                continue
        except Exception:
            pass

    manifest  = load_manifest(nb_dir)
    sec_pages = build_section_pages(manifest, nb_name)

    if not sec_pages:
        print(f"[Skip — no sections to cluster]: {nb_name}")
        total_skipped += 1
        continue

    print(f"Notebook: {nb_name}  ({len(sec_pages)} section(s) to cluster)")
    result_config = {}

    for sec_key, page_names in sorted(sec_pages.items()):
        sec_display = sec_key.split("/")[-1] if sec_key else nb_name
        print(f"  Section: {sec_display}  ({len(page_names)} pages)", end="", flush=True)

        # Build page list with summary snippets
        lines = []
        for name in page_names:
            snippet = get_summary_snippet(nb_name, sec_key, name)
            if snippet:
                lines.append(f"{name} | {snippet}")
            else:
                lines.append(name)

        prompt = CLUSTER_PROMPT.format(
            section   = sec_display,
            notebook  = nb_name,
            page_list = "\n".join(lines),
        )

        response = call_gemini(client, prompt)
        groups   = parse_llm_json(response)

        if groups is None:
            print(f"  [Error — could not parse LLM response]")
            # Fall back: put all pages in scaffold
            result_config[sec_display] = {"_pages_to_group": page_names}
            total_errors += 1
            continue

        # Validate: check all page names are accounted for
        assigned  = {p for members in groups.values() for p in members}
        missing   = [p for p in page_names if p not in assigned]
        extra     = [p for p in assigned if p not in page_names]

        if extra:
            # LLM hallucinated page names — remove them
            groups = {
                g: [p for p in members if p in set(page_names)]
                for g, members in groups.items()
            }

        if missing:
            groups.setdefault("Other", []).extend(missing)

        # Remove empty groups
        groups = {g: m for g, m in groups.items() if m}

        n_groups = len(groups)
        print(f" → {n_groups} group(s): {', '.join(list(groups)[:4])}{'...' if n_groups > 4 else ''}")
        result_config[sec_display] = groups

    # Write rollup_groups.json
    out_path.write_text(json.dumps(result_config, indent=2), encoding="utf-8")
    print(f"  ✓ Written: {out_path}\n")
    total_generated += 1

print("=" * 60)
print(f"  CLUSTERING COMPLETE")
print(f"  Generated : {total_generated}")
print(f"  Skipped   : {total_skipped}")
print(f"  Errors    : {total_errors}")
print("=" * 60)
print()
print("Next steps:")
print("  1. Review the generated rollup_groups.json files in Dropbox")
print("  2. Correct any wrong groupings")
print("  3. SCP to Cerebro:")
for nb in selected:
    print(f"       scp '{DROPBOX_BASE}/01_Raw_Audit/{nb}/rollup_groups.json' \\")
    print(f"           lab-user@10.254.254.48:Onenote-archivist/onenote_audit/01_Raw_Audit/{nb}/")
print("  4. On Cerebro: python3 summarize_rollups.py --force")
