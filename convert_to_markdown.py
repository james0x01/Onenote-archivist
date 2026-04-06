import os
import re
import base64
import requests
from pathlib import Path
from bs4 import BeautifulSoup
from markdownify import MarkdownConverter

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------

OLLAMA_HOST    = "http://10.254.254.48:11434"
VISION_MODEL   = "gemma3:12b"
DESCRIBE_IMAGES = True   # Set to False for a fast run without image descriptions

RAW_DIR = Path("onenote_audit/01_Raw_Audit")
MD_DIR  = Path("onenote_audit/02_Markdown")

# ---------------------------------------------------------------------------
# OLLAMA VISION
# ---------------------------------------------------------------------------

def describe_image(image_path):
    """Send a local image to Ollama and return a plain-text description."""
    try:
        with open(image_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("utf-8")

        resp = requests.post(
            f"{OLLAMA_HOST}/api/generate",
            json={
                "model": VISION_MODEL,
                "prompt": (
                    "Describe this image in detail for use in a technical document. "
                    "If it contains text, transcribe it exactly. "
                    "If it shows a diagram, chart, network map, or technical drawing, "
                    "explain precisely what it depicts. Be thorough."
                ),
                "images": [b64],
                "stream": False,
            },
            timeout=180,
        )
        if resp.status_code == 200:
            return resp.json().get("response", "").strip()
        print(f"      [Ollama {resp.status_code}]: {resp.text[:200]}")
        return None
    except Exception as e:
        print(f"      [Image description error]: {e}")
        return None


# ---------------------------------------------------------------------------
# INDENTATION HELPER
# ---------------------------------------------------------------------------

def get_indent_level(style_str):
    """
    Parse a CSS style string and return an integer indent level.
    OneNote uses margin-left (in pt or px) to show nested indentation.
    Examples: margin-left:36pt → level 1,  margin-left:72pt → level 2
    """
    if not style_str:
        return 0
    match = re.search(r'margin-left\s*:\s*([\d.]+)(pt|px)', style_str)
    if not match:
        return 0
    value = float(match.group(1))
    unit  = match.group(2)
    base  = 36 if unit == "pt" else 40   # approx pixels per indent level
    return max(0, round(value / base))


# ---------------------------------------------------------------------------
# CUSTOM MARKDOWN CONVERTER
# ---------------------------------------------------------------------------

class OneNoteConverter(MarkdownConverter):
    """
    Extends markdownify to handle two OneNote-specific cases:
      1. <p style="margin-left:..."> — converts to indented markdown
      2. <img src="media/...">       — local image ref + optional Ollama description
    Everything else (headings, bold, italic, tables, lists, code) is handled
    by the base markdownify library.
    """

    def __init__(self, page_dir, **options):
        super().__init__(**options)
        self.page_dir = Path(page_dir)   # absolute path to the raw page folder

    # --- Indentation ---
    def convert_p(self, el, text, convert_as_inline=False, **kwargs):
        if not text.strip():
            return ""
        level  = get_indent_level(el.get("style", ""))
        indent = "    " * level           # 4 spaces per level
        # Apply indent to every line in case the paragraph wraps
        indented = "\n".join(
            indent + line if line.strip() else line
            for line in text.splitlines()
        )
        return f"{indented}\n\n"

    # --- Images ---
    def convert_img(self, el, text, convert_as_inline=False, **kwargs):
        src = el.get("src", "")
        alt = el.get("alt", "Image")

        if not src or src.startswith("http"):
            return ""   # remote URLs were already handled in the archive step

        abs_path = (self.page_dir / src).resolve()
        img_ref  = f"\n![{alt}]({abs_path})\n"

        if DESCRIBE_IMAGES and abs_path.exists():
            print(f"      Describing: {src} ...")
            description = describe_image(abs_path)
            if description:
                quoted = "\n".join(f"> {line}" for line in description.splitlines())
                return f"{img_ref}\n> **[Image Description]**\n{quoted}\n\n"

        return f"{img_ref}\n"


# ---------------------------------------------------------------------------
# PAGE CONVERTER
# ---------------------------------------------------------------------------

def convert_page(html_file, page_dir, md_file, metadata):
    """Read index.html, convert to Markdown with YAML frontmatter, write .md file."""

    with open(html_file, "r", encoding="utf-8") as f:
        html = f.read()

    soup = BeautifulSoup(html, "html.parser")

    # Pull creation date from OneNote's <meta name="created"> if present
    meta_tag = soup.find("meta", attrs={"name": "created"})
    if meta_tag:
        metadata["created"] = meta_tag.get("content", "")

    # --- YAML frontmatter ---
    fm_lines = ["---"]
    for key, val in metadata.items():
        if val:
            # Escape any double quotes inside values
            safe_val = str(val).replace('"', '\\"')
            fm_lines.append(f'{key}: "{safe_val}"')
    fm_lines.append("---\n")
    frontmatter = "\n".join(fm_lines)

    # --- Convert HTML body to Markdown ---
    body      = soup.find("body") or soup
    converter = OneNoteConverter(page_dir=page_dir, heading_style="ATX")
    md_body   = converter.convert(str(body))

    # Collapse 3+ consecutive blank lines down to 2
    md_body = re.sub(r"\n{3,}", "\n\n", md_body).strip()

    # --- Write output ---
    os.makedirs(md_file.parent, exist_ok=True)
    with open(md_file, "w", encoding="utf-8") as f:
        f.write(frontmatter + "\n" + md_body + "\n")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

print("=" * 60)
print("  OneNote HTML → Markdown Converter")
print("=" * 60)
print(f"  Source : {RAW_DIR.resolve()}")
print(f"  Output : {MD_DIR.resolve()}")
print(f"  Ollama : {OLLAMA_HOST}")
print(f"  Model  : {VISION_MODEL}")
print(f"  Images : {'Describe via Ollama' if DESCRIBE_IMAGES else 'Reference only (fast mode)'}")
print()

# --- Test Ollama connection ---
print("Testing Ollama connection...")
try:
    r = requests.get(f"{OLLAMA_HOST}/api/tags", timeout=10)
    if r.status_code == 200:
        available = [m["name"] for m in r.json().get("models", [])]
        model_found = any(VISION_MODEL in m for m in available)
        if DESCRIBE_IMAGES and not model_found:
            print(f"  WARNING: {VISION_MODEL} not found on Ollama server.")
            print(f"  Available models: {available}")
            print(f"  Disabling image descriptions.\n")
            DESCRIBE_IMAGES = False
        else:
            status = f"{VISION_MODEL} ready" if DESCRIBE_IMAGES else "connected"
            print(f"  Ollama {status}.\n")
    else:
        print(f"  WARNING: Ollama returned {r.status_code}. Disabling image descriptions.\n")
        DESCRIBE_IMAGES = False
except Exception as e:
    print(f"  WARNING: Cannot reach Ollama at {OLLAMA_HOST}: {e}")
    print(f"  Disabling image descriptions.\n")
    DESCRIBE_IMAGES = False

# --- Walk and convert ---
total_converted = 0
total_skipped   = 0
total_errors    = 0

all_pages = sorted(RAW_DIR.rglob("index.html"))
print(f"Found {len(all_pages)} pages to process.\n")

for html_file in all_pages:
    page_dir = html_file.parent
    rel_path = page_dir.relative_to(RAW_DIR)
    parts    = rel_path.parts

    # Extract metadata from folder path
    # Structure: Notebook / [SectionGroup /] Section / PageTitle
    notebook   = parts[0] if len(parts) >= 1 else "Unknown"
    page_title = parts[-1] if len(parts) >= 1 else "Unknown"
    section    = " / ".join(parts[1:-1]) if len(parts) >= 3 else parts[1] if len(parts) == 2 else ""

    metadata = {
        "notebook" : notebook,
        "section"  : section,
        "title"    : page_title,
        "created"  : "",         # filled in by convert_page if available
        "source"   : str(html_file.resolve()),
    }

    # Output: 02_Markdown/Notebook/Section/PageTitle.md  (flat .md, no subfolder)
    md_file = MD_DIR / rel_path.parent / f"{page_title}.md"

    # --- Resume: skip already-converted pages ---
    if md_file.exists():
        print(f"  [Skip] {notebook} > {section} > {page_title}")
        total_skipped += 1
        continue

    print(f"  Converting: {notebook} > {section} > {page_title}")

    try:
        convert_page(html_file, page_dir, md_file, metadata)
        total_converted += 1
    except Exception as e:
        print(f"    [ERROR]: {e}")
        total_errors += 1

print(f"\n{'=' * 60}")
print(f"  CONVERSION COMPLETE")
print(f"{'=' * 60}")
print(f"  Converted : {total_converted}")
print(f"  Skipped   : {total_skipped}  (already done)")
print(f"  Errors    : {total_errors}")
print(f"  Output    : {MD_DIR.resolve()}")
