#!/usr/bin/env python3
"""
summarize_pages.py — Phase 1 of the OneNote summarisation pipeline.

Reads pages from 02_Markdown, summarises each using a local Ollama LLM,
and writes results to 03_Summaries with the same folder hierarchy.

Obsidian vault root: onenote_audit/
  Each summary links back to its 02_Markdown source and to any original
  attachment files in 01_Raw_Audit so you can navigate all three layers.

TODO: Phase 2 — summarize_rollups.py will read these page summaries and
      generate _section_summary.md and _notebook_summary.md rollups.
"""

import os
import re
import sys
import time
import argparse
import requests
from pathlib import Path
from datetime import datetime

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------

OLLAMA_HOST    = "http://10.254.254.48:11434"
OLLAMA_MODEL   = "qwen2.5:14b"   # run `ollama list` on the server to confirm
OLLAMA_TIMEOUT = 300             # seconds — CPU inference can be slow

VAULT_DIR = Path("onenote_audit")            # Obsidian vault root
MD_DIR    = VAULT_DIR / "02_Markdown"
RAW_DIR   = VAULT_DIR / "01_Raw_Audit"
SUM_DIR   = VAULT_DIR / "03_Summaries"

# ---------------------------------------------------------------------------
# PROMPTS
# ---------------------------------------------------------------------------

SYNTHESIS_PROMPT = """\
You are summarizing a page of personal notes from a OneNote knowledge base.

IMPORTANT — indentation (4 spaces per level) shows hierarchical relationships:
- Indented items belong to and are sub-topics of the un-indented item above them
- Example: a book title with indented lines below = chapter notes for that book
- Example: a topic with indented sub-items = related details or examples

Write a thorough synthesis that:
1. Preserves hierarchical relationships — make clear what belongs to what
2. Captures all key facts, decisions, names, dates, and action items
3. Accurately represents lists, tables, technical specs, or structured data
4. Includes relevant context from any image descriptions present
5. Is written in clear prose that mirrors the original structure

At the end, on a new line starting with "TAGS:", suggest 3-8 short lowercase \
tags that best describe the content.
Example: TAGS: networking, security, f5, competitive-analysis

Notes:
{content}
"""

# Used for Yes / No / Maybe pages inside a Candidates section
CANDIDATES_PROMPT = """\
The following is a page of interview candidate names listed under the \
category "{page_title}".
Extract only the candidate names exactly as they appear.
Return a simple markdown list of names — no analysis, no commentary.

Notes:
{content}
"""

# ---------------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------------

def safe_log(msg):
    """Log to screen and file if the log file is open, otherwise just print."""
    print(msg)
    try:
        _log_file.write(msg + "\n")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# TELEMETRY HELPER
# ---------------------------------------------------------------------------

def format_eta(seconds):
    """Format a duration in seconds as a human-readable string."""
    if seconds < 60:
        return f"{int(seconds)}s"
    elif seconds < 3600:
        return f"{int(seconds // 60)}m {int(seconds % 60)}s"
    else:
        hours   = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        return f"{hours}h {minutes}m"


# ---------------------------------------------------------------------------
# FRONTMATTER + TAG HELPERS
# ---------------------------------------------------------------------------

def parse_frontmatter(text):
    """Return (meta_dict, body_text) from a markdown file with YAML frontmatter."""
    meta = {}
    body = text
    if text.startswith("---"):
        end = text.find("---", 3)
        if end != -1:
            fm = text[3:end]
            for line in fm.splitlines():
                if ":" in line:
                    key, _, val = line.partition(":")
                    meta[key.strip()] = val.strip().strip('"')
            body = text[end + 3:].strip()
    return meta, body


def parse_tags_from_response(response):
    """
    Split the LLM response into (summary_text, tags_list).
    The model is asked to append a line like: TAGS: foo, bar, baz
    """
    lines       = response.strip().splitlines()
    tags        = []
    body_lines  = []
    for line in lines:
        stripped = line.strip()
        if stripped.upper().startswith("TAGS:"):
            raw_tags = stripped.split(":", 1)[1]
            tags = [
                t.strip().lower().replace(" ", "-")
                for t in raw_tags.split(",")
                if t.strip()
            ]
        else:
            body_lines.append(line)
    return "\n".join(body_lines).strip(), tags


# ---------------------------------------------------------------------------
# ATTACHMENT HELPER
# ---------------------------------------------------------------------------

def find_attachments(md_file):
    """
    Given a path inside 02_Markdown, return a list of attachment file paths
    from the corresponding 01_Raw_Audit page folder.
    """
    # md_file: onenote_audit/02_Markdown/Notebook/.../PageTitle.md
    # raw dir: onenote_audit/01_Raw_Audit/Notebook/.../PageTitle/attachments/
    rel        = md_file.relative_to(MD_DIR)          # Notebook/.../PageTitle.md
    page_stem  = rel.with_suffix("")                   # Notebook/.../PageTitle
    attach_dir = RAW_DIR / page_stem / "attachments"
    if not attach_dir.exists():
        return []
    return sorted(attach_dir.iterdir())


def wiki_link(file_path, display_name=None):
    """
    Build an Obsidian wiki-link using a path relative to the vault root.
    Markdown files omit the extension; other files keep it.
    """
    rel  = file_path.relative_to(VAULT_DIR)
    name = display_name or file_path.name
    if file_path.suffix.lower() == ".md":
        rel = rel.with_suffix("")
    return f"[[{rel}|{name}]]"


# ---------------------------------------------------------------------------
# OLLAMA
# ---------------------------------------------------------------------------

def call_ollama(prompt):
    """Send a prompt to Ollama and return the response string."""
    try:
        resp = requests.post(
            f"{OLLAMA_HOST}/api/generate",
            json={"model": OLLAMA_MODEL, "prompt": prompt, "stream": False},
            timeout=OLLAMA_TIMEOUT,
        )
        resp.raise_for_status()
        return resp.json()["response"].strip()
    except requests.exceptions.Timeout:
        safe_log(f"    [Summarisation error]: Ollama timed out after {OLLAMA_TIMEOUT}s")
        return None
    except Exception as e:
        safe_log(f"    [Summarisation error]: {e}")
        return None


def summarise_page(body, section, page_title):
    """Choose the right prompt and call Ollama. Returns (summary, tags)."""
    is_candidate_page = (
        section.lower().endswith("candidates") and
        page_title.lower() in ("yes", "no", "maybe")
    )

    if is_candidate_page:
        prompt   = CANDIDATES_PROMPT.format(page_title=page_title, content=body)
        response = call_ollama(prompt)
        if response is None:
            return None, []
        return response, []          # no tags for candidate list pages
    else:
        prompt   = SYNTHESIS_PROMPT.format(content=body)
        response = call_ollama(prompt)
        if response is None:
            return None, []
        return parse_tags_from_response(response)


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

parser = argparse.ArgumentParser(description="Summarise 02_Markdown pages with Ollama")
parser.add_argument(
    "--force", action="store_true",
    help="Re-summarise pages that already have a summary file"
)
args = parser.parse_args()

print("=" * 60)
print("  OneNote Markdown → Page Summaries")
print("=" * 60)
print(f"  Source : {MD_DIR.resolve()}")
print(f"  Output : {SUM_DIR.resolve()}")
print(f"  Model  : {OLLAMA_MODEL} via {OLLAMA_HOST}")
print()

# --- Log file ---
os.makedirs(str(VAULT_DIR), exist_ok=True)
log_path  = str(VAULT_DIR / f"03_Summaries-log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt")
_log_file = open(log_path, "w", encoding="utf-8", buffering=1)
print(f"Logging to: {log_path}\n")

def log(msg=""):
    print(msg)
    _log_file.write(msg + "\n")

# --- Check Ollama is reachable and model is available ---
try:
    r      = requests.get(f"{OLLAMA_HOST}/api/tags", timeout=10)
    r.raise_for_status()
    models = [m["name"] for m in r.json().get("models", [])]
    if not any(OLLAMA_MODEL in m for m in models):
        log(f"  WARNING: '{OLLAMA_MODEL}' not found in Ollama.")
        log(f"  Available: {', '.join(models)}")
        log(f"  Update OLLAMA_MODEL at the top of this script and re-run.\n")
        _log_file.close()
        sys.exit(1)
    log(f"  Ollama ready. Model: {OLLAMA_MODEL}\n")
except requests.exceptions.ConnectionError:
    log(f"  ERROR: Cannot reach Ollama at {OLLAMA_HOST}")
    log(f"  Check that Ollama is running and OLLAMA_HOST is correct.\n")
    _log_file.close()
    sys.exit(1)

# --- Force / overwrite prompt ---
force = args.force
if not force:
    answer = input("Overwrite existing summaries? (y/N): ").strip().lower()
    force  = answer == "y"
if force:
    log("  Mode: overwrite — existing summaries will be re-generated.\n")
else:
    log("  Mode: resume — existing summaries will be skipped.\n")

# --- Notebook selection ---
notebooks = sorted([d.name for d in MD_DIR.iterdir() if d.is_dir()])
print("Available notebooks:")
for i, nb in enumerate(notebooks, 1):
    print(f"  {i:>2}. {nb}")
print()
print("  a           — summarise all notebooks")
print("  s <numbers> — skip these notebooks       e.g. s 1 3 5")
print("  p <numbers> — pull only these notebooks  e.g. p 2 4")
print()

selection = input("Your choice: ").strip().lower()
all_pages = sorted(MD_DIR.rglob("*.md"))

if selection == "a" or selection == "":
    log(f"\nSummarising all {len(notebooks)} notebooks.\n")
elif selection.startswith("s "):
    skip_nums  = {int(x) for x in selection[2:].split() if x.isdigit()}
    skip_names = {notebooks[i-1] for i in skip_nums if 1 <= i <= len(notebooks)}
    all_pages  = [p for p in all_pages if p.relative_to(MD_DIR).parts[0] not in skip_names]
    log(f"\nSkipping {len(skip_nums)} notebook(s). Processing remaining pages.\n")
elif selection.startswith("p "):
    pull_nums  = {int(x) for x in selection[2:].split() if x.isdigit()}
    pull_names = {notebooks[i-1] for i in pull_nums if 1 <= i <= len(notebooks)}
    all_pages  = [p for p in all_pages if p.relative_to(MD_DIR).parts[0] in pull_names]
    log(f"\nSummarising {len(pull_nums)} notebook(s).\n")
else:
    log(f"\nUnrecognised input — summarising all notebooks.\n")

# Log selected notebooks
log("Notebooks selected for this run:")
selected_nbs = sorted({p.relative_to(MD_DIR).parts[0] for p in all_pages})
for nb in selected_nbs:
    log(f"  - {nb}")
log()
log(f"Found {len(all_pages)} pages to process.\n")

# --- Walk and summarise ---
total_summarised = 0
total_skipped    = 0
total_errors     = 0
processing_times = []   # seconds per page (summarised pages only)
page_sizes       = []   # chars per page
script_start     = time.time()

for md_file in all_pages:
    rel_path   = md_file.relative_to(MD_DIR)
    parts      = rel_path.parts
    notebook   = parts[0]
    page_title = md_file.stem
    section    = " / ".join(parts[1:-1]) if len(parts) >= 3 else (parts[1] if len(parts) == 2 else "")

    sum_file = SUM_DIR / rel_path

    # --- Resume / force ---
    if sum_file.exists() and not force:
        log(f"  [Skip] {notebook} > {section} > {page_title}")
        total_skipped += 1
        continue

    log(f"  Summarising: {notebook} > {section} > {page_title}")

    try:
        text       = md_file.read_text(encoding="utf-8")
        meta, body = parse_frontmatter(text)

        if not body.strip():
            log(f"    [Skip] Empty page")
            total_skipped += 1
            continue

        page_size = len(body)
        page_start = time.time()
        summary, llm_tags = summarise_page(body, section, page_title)
        elapsed   = time.time() - page_start

        if summary is None:
            total_errors += 1
            continue

        # --- Build tag list ---
        auto_tags = []
        if notebook:
            auto_tags.append(notebook.lower().replace(" ", "-"))
        if section:
            auto_tags.extend(
                p.lower().replace(" ", "-") for p in section.split(" / ")
            )
        all_tags  = auto_tags + [t for t in llm_tags if t not in auto_tags]
        tags_yaml = "[" + ", ".join(all_tags) + "]"

        # --- Attachment links ---
        attachments    = find_attachments(md_file)
        attach_section = ""
        if attachments:
            links = "\n".join(
                f"- {wiki_link(a)}" for a in attachments
            )
            attach_section = f"\n\n## Attachments\n\n{links}"

        # --- YAML frontmatter ---
        fm_lines = ["---"]
        for key in ("notebook", "section", "title", "created"):
            val = meta.get(key, "")
            if val:
                fm_lines.append(f'{key}: "{val}"')
        fm_lines.append(f"tags: {tags_yaml}")
        fm_lines.append(f'source_markdown: "{wiki_link(md_file, page_title + " (Full Notes)")}"')
        fm_lines.append(f'summarised: "{datetime.now().strftime("%Y-%m-%d")}"')
        fm_lines.append("---\n")
        frontmatter = "\n".join(fm_lines)

        # --- Write output ---
        output = f"{frontmatter}# Summary: {page_title}\n\n{summary}{attach_section}\n"
        os.makedirs(str(sum_file.parent), exist_ok=True)
        sum_file.write_text(output, encoding="utf-8")
        total_summarised += 1

        # --- Telemetry ---
        processing_times.append(elapsed)
        page_sizes.append(page_size)
        avg_time  = sum(processing_times) / len(processing_times)
        done      = total_summarised + total_skipped + total_errors
        remaining = len(all_pages) - done
        eta       = format_eta(avg_time * remaining) if remaining > 0 else "done"
        log(f"    [{elapsed:.1f}s | {page_size:,} chars | avg {avg_time:.1f}s/page | ETA {eta}]")

    except Exception as e:
        log(f"    [ERROR]: {e}")
        total_errors += 1

total_elapsed = time.time() - script_start

log(f"\n{'=' * 60}")
log(f"  SUMMARISATION COMPLETE")
log(f"{'=' * 60}")
log(f"  Summarised : {total_summarised}")
log(f"  Skipped    : {total_skipped}  (already done or empty)")
log(f"  Errors     : {total_errors}")
log(f"  Total time : {format_eta(total_elapsed)}")
if processing_times:
    log(f"  Avg/page   : {format_eta(sum(processing_times) / len(processing_times))}")
    log(f"  Fastest    : {format_eta(min(processing_times))}  ({min(page_sizes):,} chars)")
    log(f"  Slowest    : {format_eta(max(processing_times))}  ({max(page_sizes):,} chars)")
log(f"  Output     : {SUM_DIR.resolve()}")
_log_file.close()
