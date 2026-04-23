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
import json
import time
import argparse
import requests
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
from google import genai

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------

load_dotenv()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

OLLAMA_HOST    = "http://10.254.254.48:11434"
OLLAMA_MODEL   = "qwen2.5:32b"   # default; overridden by menu selection at runtime
OLLAMA_TIMEOUT = 600             # seconds — 10 minutes for local LLM

GEMINI_MODELS = {
    "g": ("gemini-2.5-flash", "gemini-2.5-flash  (fast, cloud)"),
    "f": ("gemini-2.5-pro",   "gemini-2.5-pro    (more capacity, cloud)"),
}
GEMINI_CALL_DELAY = 2   # seconds between Gemini calls

# Display labels for known Ollama models — unlisted models show name only
OLLAMA_MODEL_LABELS = {
    "qwen2.5:14b":          "qwen2.5:14b         (14B, balanced)",
    "qwen2.5:32b":          "qwen2.5:32b         (32B, higher quality — default)",
    "llama4:latest":        "llama4              (108B, powerful, very slow on CPU)",
    "mistral-large:latest": "mistral-large       (123B, strong reasoning, slow)",
    "gemma3:12b":           "gemma3:12b          (12B, fast, lighter)",
    "phi4-mini:latest":     "phi4-mini           (3.8B, fastest, basic quality)",
}

_llm_backend    = "ollama"   # "ollama" or "gemini" — set by menu
_llm_model_name = OLLAMA_MODEL
_gemini_client  = None

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

# Used for Yes / No / Maybe / In-process pages inside a Candidates section
CANDIDATES_PROMPT = """\
The following is a list of interview candidates grouped by hiring status.
Extract the candidate names exactly as they appear, preserving their groupings.
Return a markdown list with each status as a ## heading, and names as bullet points beneath it.
Do not add analysis, commentary, or change any names.

Example output format:
## Yes
- Jane Smith
- John Doe

## No
- Bob Jones

## Maybe
- Alice Brown

Notes:
{content}
"""

# Used to generate a consolidated roster combining all status pages for a section
CANDIDATES_ROSTER_PROMPT = """\
The following contains candidate names from multiple hiring status lists for the same team.
Each section is labeled with its status (Yes, No, Maybe, In-process, etc.).
Combine them into a single clean roster with each status as a ## heading and names as bullet points.
Preserve all names exactly. Do not add analysis or commentary.

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
            json={"model": _llm_model_name, "prompt": prompt, "stream": False},
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


def call_gemini(prompt):
    """Send a prompt to Gemini and return the response string."""
    time.sleep(GEMINI_CALL_DELAY)
    wait = 20
    for attempt in range(3):
        try:
            response = _gemini_client.models.generate_content(
                model=_llm_model_name,
                contents=prompt,
            )
            return response.text.strip()
        except Exception as e:
            err = str(e)
            if "503" in err and attempt < 2:
                safe_log(f"    [Retry {attempt+1}/3 in {wait}s]: {e}")
                time.sleep(wait)
                wait = min(wait + 20, 60)
            else:
                safe_log(f"    [Summarisation error]: {e}")
                return None
    return None


def call_llm(prompt):
    """Dispatch to the selected LLM backend."""
    if _llm_backend == "gemini":
        return call_gemini(prompt)
    return call_ollama(prompt)


def summarise_page(body, section, page_title):
    """Choose the right prompt and call the LLM. Returns (summary, tags)."""
    CANDIDATE_STATUS_PAGES = {"yes", "no", "maybe", "in-process", "in process", "pipeline"}
    is_candidate_page = (
        section.lower().endswith("candidates") and
        page_title.lower() in CANDIDATE_STATUS_PAGES
    )

    if is_candidate_page:
        prompt   = CANDIDATES_PROMPT.format(page_title=page_title, content=body)
        response = call_llm(prompt)
        if response is None:
            return None, []
        return response, []          # no tags for candidate list pages
    else:
        prompt   = SYNTHESIS_PROMPT.format(content=body)
        response = call_llm(prompt)
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
print()

# --- Log file ---
os.makedirs(str(VAULT_DIR), exist_ok=True)
log_path  = str(VAULT_DIR / f"03_Summaries-log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt")
_log_file = open(log_path, "w", encoding="utf-8", buffering=1)
print(f"Logging to: {log_path}\n")

def log(msg=""):
    print(msg)
    _log_file.write(msg + "\n")

# --- LLM model selection ---
# Fetch available Ollama models for the menu
_ollama_models = []
try:
    r = requests.get(f"{OLLAMA_HOST}/api/tags", timeout=10)
    r.raise_for_status()
    _ollama_models = sorted(m["name"] for m in r.json().get("models", []))
except Exception:
    pass

print("LLM model for summarisation:")
for key, (_, label) in GEMINI_MODELS.items():
    print(f"  {key}  — Gemini: {label}")
if _ollama_models:
    for idx, name in enumerate(_ollama_models, 1):
        label = OLLAMA_MODEL_LABELS.get(name, name)
        print(f"  {idx}  — Ollama: {label}")
else:
    print(f"  (Ollama unreachable at {OLLAMA_HOST})")
print()

llm_choice = input(f"Model choice [{OLLAMA_MODEL}]: ").strip().lower()
print()

if llm_choice in GEMINI_MODELS:
    _llm_backend    = "gemini"
    _llm_model_name, model_label = GEMINI_MODELS[llm_choice]
    if not GEMINI_API_KEY:
        print("  ERROR: GEMINI_API_KEY not found in .env. Cannot use Gemini.")
        _log_file.close()
        sys.exit(1)
    _gemini_client = genai.Client(api_key=GEMINI_API_KEY)
    log(f"  LLM: Gemini — {model_label}\n")
elif llm_choice.isdigit() and 1 <= int(llm_choice) <= len(_ollama_models):
    _llm_backend    = "ollama"
    _llm_model_name = _ollama_models[int(llm_choice) - 1]
    log(f"  LLM: Ollama — {_llm_model_name}\n")
else:
    # Default — use OLLAMA_MODEL
    _llm_backend    = "ollama"
    _llm_model_name = OLLAMA_MODEL
    if _ollama_models and not any(_llm_model_name in m for m in _ollama_models):
        log(f"  WARNING: '{_llm_model_name}' not found in Ollama.")
        log(f"  Available: {', '.join(_ollama_models)}\n")
        _log_file.close()
        sys.exit(1)
    log(f"  LLM: Ollama — {_llm_model_name}\n")

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
total_updated    = 0
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

    # Read source file first — needed for timestamp check and summarisation
    try:
        text       = md_file.read_text(encoding="utf-8")
        meta, body = parse_frontmatter(text)
    except Exception as e:
        log(f"    [ERROR reading source]: {e}")
        total_errors += 1
        continue

    if not body.strip():
        log(f"  [Skip] Empty page: {notebook} > {section} > {page_title}")
        total_skipped += 1
        continue

    source_ts = meta.get("lastModifiedDateTime", "")

    # --- Resume / force ---
    is_update = False
    if sum_file.exists() and not force:
        if source_ts:
            try:
                existing_meta, _ = parse_frontmatter(sum_file.read_text(encoding="utf-8"))
                if existing_meta.get("lastModifiedDateTime", "") == source_ts:
                    log(f"  [Skip] {notebook} > {section} > {page_title}")
                    total_skipped += 1
                    continue
                # Timestamp changed — re-summarise
                is_update = True
                log(f"  [Update] {notebook} > {section} > {page_title}")
            except Exception:
                log(f"  [Update] {notebook} > {section} > {page_title}")
                is_update = True
        else:
            # No timestamp in source — fall back to old behaviour
            log(f"  [Skip] {notebook} > {section} > {page_title}")
            total_skipped += 1
            continue
    else:
        log(f"  Summarising: {notebook} > {section} > {page_title}")

    try:

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
        for key in ("notebook", "section", "title", "created", "lastModifiedDateTime"):
            val = meta.get(key, "")
            if val:
                fm_lines.append(f'{key}: "{val}"')
        fm_lines.append(f"tags: {tags_yaml}")
        fm_lines.append(f'source_markdown: "{wiki_link(md_file, page_title + " (Full Notes)")}"')
        fm_lines.append(f'summarised: "{datetime.now().strftime("%Y-%m-%d")}"')
        fm_lines.append(f'summarised_by: "{_llm_model_name}"')
        fm_lines.append("---\n")
        frontmatter = "\n".join(fm_lines)

        # --- Write output ---
        output = f"{frontmatter}# Summary: {page_title}\n\n{summary}{attach_section}\n"
        os.makedirs(str(sum_file.parent), exist_ok=True)
        sum_file.write_text(output, encoding="utf-8")
        if is_update:
            total_updated += 1
        else:
            total_summarised += 1

        # --- Telemetry ---
        processing_times.append(elapsed)
        page_sizes.append(page_size)
        avg_time  = sum(processing_times) / len(processing_times)
        done      = total_summarised + total_updated + total_skipped + total_errors
        remaining = len(all_pages) - done
        eta       = format_eta(avg_time * remaining) if remaining > 0 else "done"
        log(f"    [{elapsed:.1f}s | {page_size:,} chars | avg {avg_time:.1f}s/page | ETA {eta}]")

    except Exception as e:
        log(f"    [ERROR]: {e}")
        total_errors += 1

total_elapsed = time.time() - script_start

# ---------------------------------------------------------------------------
# CONSOLIDATED CANDIDATE ROSTERS
# Uses page level/order from manifest to group candidates under their
# status header (Yes/No/Maybe/In-process) exactly as they appear in OneNote.
# Falls back to listing all candidates alphabetically if manifest has no
# hierarchy data (e.g. pages pulled before this fix was deployed).
# ---------------------------------------------------------------------------

CANDIDATE_STATUS_PAGES = {"yes", "no", "maybe", "in-process", "in process", "pipeline"}

def load_manifest_for(notebook):
    """Load manifest.json for a notebook from 01_Raw_Audit."""
    path = RAW_DIR / notebook / "manifest.json"
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {"pages": {}}

# Find all Candidates sections in selected notebooks
roster_notebooks = set()
for md_file in MD_DIR.rglob("*.md"):
    parts = md_file.relative_to(MD_DIR).parts
    if (len(parts) >= 2
            and md_file.parent.name.lower().endswith("candidates")
            and parts[0] in selected_nbs):
        roster_notebooks.add(parts[0])

if roster_notebooks:
    log(f"\nGenerating candidate rosters for: {', '.join(sorted(roster_notebooks))}")

for notebook in sorted(roster_notebooks):
    candidates_dir = MD_DIR / notebook / "Candidates"
    if not candidates_dir.exists():
        # Try case-insensitive match
        matches = [d for d in (MD_DIR / notebook).iterdir()
                   if d.is_dir() and d.name.lower() == "candidates"]
        if not matches:
            continue
        candidates_dir = matches[0]

    roster_file = SUM_DIR / notebook / candidates_dir.name / "_Candidates_Roster.md"
    manifest    = load_manifest_for(notebook)
    pages_meta  = manifest.get("pages", {})

    # Collect all pages in this Candidates section with their level/order
    section_key_prefix = f"Candidates/"
    candidate_pages = []
    for key, meta in pages_meta.items():
        if key.startswith(section_key_prefix):
            page_name = key[len(section_key_prefix):]
            candidate_pages.append({
                "name":  page_name,
                "level": meta.get("level", 0),
                "order": meta.get("order", 0),
            })

    # Sort by order (response index when API level/order fields are absent)
    candidate_pages.sort(key=lambda p: p["order"])

    if not candidate_pages:
        log(f"  [Roster skip] No pages found in manifest for {notebook}/Candidates")
        continue

    # -----------------------------------------------------------------------
    # Positional-divider grouping.
    #
    # The OneNote Graph API for personal accounts does not return level/order
    # fields.  However, pages are returned in their visual section order.
    # patch_manifest_hierarchy.py stores the response index as 'order', so
    # sorting by order recreates the OneNote visual sequence.
    #
    # We walk in that order: blank status pages (Yes/No/Maybe/…) act as
    # section dividers; everything that follows belongs to that group until
    # the next divider.  Pages before any status page → "Other / Unassigned".
    # Non-status, non-blank level-0 pages (templates etc.) → "Templates & Other".
    # -----------------------------------------------------------------------

    def _page_link(name):
        sp = SUM_DIR / notebook / candidates_dir.name / f"{name}.md"
        return f"- [[{name}]]" if sp.exists() else f"- {name}"

    def _is_blank(name):
        md = candidates_dir / f"{name}.md"
        if not md.exists():
            return True
        text = md.read_text(encoding="utf-8")
        if text.startswith("---"):
            end = text.find("---", 3)
            body = text[end + 3:].strip() if end != -1 else text
        else:
            body = text.strip()
        return not body

    def _is_template_header(name):
        """Pages ending in 'template' act as group headers for their sub-pages."""
        return name.lower().endswith("template") or name.lower().endswith("templates")

    def _is_version_page(name):
        """Sub-pages of templates are version pages: V1, V2, V1-AppSec, v2-General…"""
        return bool(re.match(r'^[Vv]\d', name))

    # -----------------------------------------------------------------------
    # Pass 1 — identify template headers and their IMMEDIATELY following
    # version pages (V1, V2, V1-AppSec…).  Only consecutive version-named
    # pages count; the first non-version page ends the template group.
    # This prevents candidates who happen to follow a template page from
    # being incorrectly grouped under it.
    # -----------------------------------------------------------------------
    template_groups   = []    # [{"heading": str, "members": [str]}]
    template_page_set = set() # all page names consumed by template grouping

    i = 0
    while i < len(candidate_pages):
        name = candidate_pages[i]["name"]
        if _is_template_header(name):
            versions = []
            j = i + 1
            while j < len(candidate_pages) and _is_version_page(candidate_pages[j]["name"]):
                versions.append(candidate_pages[j]["name"])
                j += 1
            template_groups.append({"heading": name, "members": versions})
            template_page_set.add(name)
            template_page_set.update(versions)
            i = j
        else:
            i += 1

    # -----------------------------------------------------------------------
    # Pass 2 — process candidates using Yes/No/Maybe blank pages as dividers,
    # skipping any pages already claimed by a template group.
    # -----------------------------------------------------------------------
    status_groups = []   # [{"heading": str, "members": [str]}]
    current_group = None
    other_pages   = []   # pages before the first status divider

    for p in candidate_pages:
        name = p["name"]
        if name in template_page_set:
            continue   # handled in pass 1

        is_status = name.lower() in CANDIDATE_STATUS_PAGES
        blank     = _is_blank(name)

        if is_status and blank:
            current_group = {"heading": name, "members": []}
            status_groups.append(current_group)
        else:
            (current_group["members"] if current_group else other_pages).append(name)

    # --- Render ---
    roster_lines = []
    for g in status_groups:
        roster_lines.append(f"\n## {g['heading']}\n")
        for m in g["members"]:
            roster_lines.append(_page_link(m))

    if other_pages:
        unassigned = [n for n in other_pages if _is_blank(n)]
        loose      = [n for n in other_pages if not _is_blank(n)]
        if unassigned:
            roster_lines.append("\n\n---\n\n## Unassigned\n")
            for n in unassigned:
                roster_lines.append(_page_link(n))
        if loose:
            roster_lines.append("\n\n---\n\n## Other\n")
            for n in loose:
                roster_lines.append(_page_link(n))

    if template_groups:
        roster_lines.append("\n\n---\n\n## Templates & Other Pages\n")
        for g in template_groups:
            roster_lines.append(f"\n### {g['heading']}\n")
            for m in g["members"]:
                roster_lines.append(_page_link(m))

    roster_content = "\n".join(roster_lines).strip()
    log(f"  Roster (positional): {notebook}/Candidates")

    fm_lines = [
        "---",
        f'notebook: "{notebook}"',
        f'title: "Candidates Roster"',
        f'summarised: "{datetime.now().strftime("%Y-%m-%d")}"',
        f'summarised_by: "{_llm_model_name}"',
        "tags: [candidates, roster]",
        "---\n",
    ]
    os.makedirs(str(roster_file.parent), exist_ok=True)
    output = "\n".join(fm_lines) + f"# Candidates Roster\n\n{roster_content}\n"
    roster_file.write_text(output, encoding="utf-8")

log(f"\n{'=' * 60}")
log(f"  SUMMARISATION COMPLETE")
log(f"{'=' * 60}")
log(f"  Summarised : {total_summarised}")
log(f"  Updated    : {total_updated}")
log(f"  Skipped    : {total_skipped}  (unchanged or empty)")
log(f"  Errors     : {total_errors}")
log(f"  Total time : {format_eta(total_elapsed)}")
if processing_times:
    log(f"  Avg/page   : {format_eta(sum(processing_times) / len(processing_times))}")
    log(f"  Fastest    : {format_eta(min(processing_times))}  ({min(page_sizes):,} chars)")
    log(f"  Slowest    : {format_eta(max(processing_times))}  ({max(page_sizes):,} chars)")
log(f"  Output     : {SUM_DIR.resolve()}")
_log_file.close()
