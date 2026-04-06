import os
import re
import csv
import requests
import PIL.Image
import google.generativeai as genai
from pathlib import Path
from bs4 import BeautifulSoup
from markdownify import MarkdownConverter
from dotenv import load_dotenv
from docx import Document
from pptx import Presentation
from pypdf import PdfReader
import openpyxl

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------

load_dotenv()
GEMINI_API_KEY  = os.getenv("GEMINI_API_KEY")
VISION_MODEL    = "gemini-1.5-flash"
DESCRIBE_IMAGES = True   # Set to False for a fast run without image descriptions

RAW_DIR = Path("onenote_audit/01_Raw_Audit")
MD_DIR  = Path("onenote_audit/02_Markdown")

# ---------------------------------------------------------------------------
# GEMINI VISION
# ---------------------------------------------------------------------------

def describe_image(image_path):
    """Send a local image to Gemini and return a plain-text description."""
    try:
        img = PIL.Image.open(image_path)
        response = genai.GenerativeModel(VISION_MODEL).generate_content([
            (
                "Describe this image in detail for use in a technical document. "
                "If it contains text, transcribe it exactly. "
                "If it shows a diagram, chart, network map, or technical drawing, "
                "explain precisely what it depicts. Be thorough."
            ),
            img,
        ])
        return response.text.strip()
    except Exception as e:
        print(f"      [Image description error]: {e}")
        return None


# ---------------------------------------------------------------------------
# ATTACHMENT TEXT EXTRACTION
# ---------------------------------------------------------------------------

def extract_attachment_text(file_path):
    """
    Extract readable text from a local attachment file.
    Supports: .docx, .pptx, .pdf, .xlsx, .csv, .txt
    Returns a markdown-formatted string of the content.
    """
    suffix = Path(file_path).suffix.lower()
    try:
        if suffix == ".docx":
            doc   = Document(file_path)
            lines = [p.text for p in doc.paragraphs if p.text.strip()]
            return "\n\n".join(lines)

        elif suffix == ".pptx":
            prs   = Presentation(file_path)
            lines = []
            for i, slide in enumerate(prs.slides, 1):
                lines.append(f"### Slide {i}")
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text.strip():
                        lines.append(shape.text.strip())
            return "\n\n".join(lines)

        elif suffix == ".pdf":
            reader = PdfReader(file_path)
            lines  = []
            for i, page in enumerate(reader.pages, 1):
                text = page.extract_text()
                if text and text.strip():
                    lines.append(f"### Page {i}\n{text.strip()}")
            return "\n\n".join(lines)

        elif suffix == ".xlsx":
            wb    = openpyxl.load_workbook(file_path, data_only=True)
            lines = []
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                lines.append(f"### Sheet: {sheet_name}")
                for row in ws.iter_rows(values_only=True):
                    row_text = " | ".join(
                        str(cell) if cell is not None else "" for cell in row
                    )
                    if row_text.strip(" |"):
                        lines.append(row_text)
            return "\n\n".join(lines)

        elif suffix == ".csv":
            with open(file_path, "r", encoding="utf-8", errors="replace") as f:
                rows = list(csv.reader(f))
            if not rows:
                return ""
            # Format as a markdown table
            header = "| " + " | ".join(rows[0]) + " |"
            sep    = "| " + " | ".join(["---"] * len(rows[0])) + " |"
            body   = "\n".join("| " + " | ".join(r) + " |" for r in rows[1:])
            return "\n".join([header, sep, body])

        elif suffix == ".txt":
            with open(file_path, "r", encoding="utf-8", errors="replace") as f:
                return f.read().strip()

        else:
            return None  # unsupported type — will fall back to reference only

    except Exception as e:
        print(f"      [Extraction error for {Path(file_path).name}]: {e}")
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

    # --- Attachments (Word, PowerPoint, PDF, Excel, CSV, TXT) ---
    def convert_object(self, el, text, convert_as_inline=False, **kwargs):
        filename = el.get("data-attachment", "")
        data_src = el.get("data", "")

        if not filename:
            return ""

        # Resolve the local path — archive script saves to attachments/filename
        attach_path = (self.page_dir / data_src).resolve() if data_src else None

        header = f"\n---\n📎 **Attachment: {filename}**\n"

        if attach_path and attach_path.exists():
            content = extract_attachment_text(attach_path)
            if content:
                print(f"      Extracting: {filename}")
                return f"{header}\n{content}\n\n---\n"
            else:
                # File type not supported for extraction — reference only
                return f"{header}*(Content extraction not supported for this file type)*\n\n---\n"
        else:
            # File wasn't downloaded or doesn't exist
            return f"{header}*(File not found in attachments folder)*\n\n---\n"

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
print(f"  Model  : {VISION_MODEL} (Gemini)")
print(f"  Images : {'Describe via Gemini' if DESCRIBE_IMAGES else 'Reference only (fast mode)'}")
print()

# --- Initialise Gemini ---
if DESCRIBE_IMAGES:
    if not GEMINI_API_KEY:
        print("  WARNING: GEMINI_API_KEY not found in .env. Disabling image descriptions.\n")
        DESCRIBE_IMAGES = False
    else:
        try:
            genai.configure(api_key=GEMINI_API_KEY)
            # Quick test call to confirm the key and model are valid
            genai.GenerativeModel(VISION_MODEL).generate_content("hello")
            print(f"  Gemini ready. Model: {VISION_MODEL}\n")
        except Exception as e:
            print(f"  WARNING: Gemini initialisation failed: {e}")
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
