#!/usr/bin/env python3
"""
PDF Table Extractor — powered by Claude Vision (claude-opus-4-8)

Extracts ALL tables from very long PDF documents — including tables pasted as
images/screenshots that text-based parsers (pdfplumber, camelot) cannot see — by
rendering every page as an image and using Claude's multimodal understanding.

For each table it captures, Azure Document Intelligence style:
  - title / caption
  - footnotes, source lines, and notes
  - section / context label
  - column headers and all data rows (merged cells preserved)
  - page number, table index, and cross-page continuation links

Outputs, into a user-specified folder:
  - one CSV per table (with a metadata header block)
  - one combined CSV per PDF (every table, long format)
  - one structured JSON manifest per PDF (full Azure-DI-like dump)

Usage:
    pip install anthropic pymupdf
    export ANTHROPIC_API_KEY=sk-ant-...
    python extract_pdf_tables.py [PDF_PATH_OR_FOLDER] [OUTPUT_FOLDER]

If the two paths are omitted, the script prompts for them interactively.
"""

import os
import sys
import json
import time
import base64
import csv
import re
from pathlib import Path
from typing import Optional

import anthropic

try:
    import fitz  # pymupdf
except ImportError:
    print("Installing pymupdf (required for PDF rendering)...")
    os.system(f"{sys.executable} -m pip install pymupdf -q")
    import fitz


MODEL = "claude-opus-4-8"

# Claude's vision pipeline downsamples images to roughly 1.15 megapixels /
# ~1568 px on the long edge. Rendering far above that just wastes tokens and
# time without improving what the model actually sees. We target a long edge a
# touch above that so the page is crisp, then let Claude downsample cleanly.
TARGET_LONG_EDGE_PX = 1540
MIN_DPI = 110
MAX_DPI = 240

MAX_RETRIES = 5

# ---------------------------------------------------------------------------
# Structured-output tool schema. Forcing Claude to emit data through this tool
# guarantees well-formed structure (no brittle regex/JSON-from-prose parsing).
# ---------------------------------------------------------------------------
RECORD_TOOL = {
    "name": "record_tables",
    "description": (
        "Record every table found on the page as structured data. Call this "
        "exactly once per page, passing all tables found (or an empty list if "
        "there are none)."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "tables": {
                "type": "array",
                "description": "All tables found on this page, in reading order.",
                "items": {
                    "type": "object",
                    "properties": {
                        "table_index": {
                            "type": "integer",
                            "description": "1-based order of appearance on this page.",
                        },
                        "title": {
                            "type": ["string", "null"],
                            "description": "Table title/caption, e.g. 'Table 1: Revenue'. null if none.",
                        },
                        "context_label": {
                            "type": ["string", "null"],
                            "description": "Section heading that contextualizes the table, e.g. '3.2 Results'. null if none.",
                        },
                        "headers": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "Column header names, in order. Empty if the table has no header row.",
                        },
                        "rows": {
                            "type": "array",
                            "items": {"type": "array", "items": {"type": "string"}},
                            "description": "Data rows; each is a list of cell values in column order. Repeat values across merged/spanned cells. Use '' for empty cells.",
                        },
                        "footnotes": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "Footnotes, notes, and source lines below the table.",
                        },
                        "continued_from_previous": {
                            "type": "boolean",
                            "description": "True if this table appears to continue from the previous page (e.g. starts mid-data with no/partial header).",
                        },
                        "continues_on_next": {
                            "type": "boolean",
                            "description": "True if this table appears to continue onto the next page (e.g. runs to the bottom edge, no closing/total row).",
                        },
                    },
                    "required": [
                        "table_index", "title", "context_label", "headers",
                        "rows", "footnotes", "continued_from_previous",
                        "continues_on_next",
                    ],
                },
            }
        },
        "required": ["tables"],
    },
}

EXTRACTION_PROMPT = """You are an expert document-analysis system specializing in table extraction, comparable to Azure Document Intelligence.

Examine this PDF page image and extract EVERY table present. Critically, this includes tables that are embedded as images, screenshots, or scans — transcribe their contents exactly as if they were native tables.

For each table, capture:
- title/caption (text directly above the table)
- context_label (the section heading it sits under)
- headers (column names, in order)
- rows (every data row; values exactly as shown — preserve commas, decimals, %, currency symbols, units, parentheses for negatives, and footnote markers)
- footnotes (notes/source lines below the table)
- continuation flags (whether the table continues from the previous page or onto the next)

Rules for maximum accuracy:
- Do NOT omit, summarize, or reorder any row or column.
- For merged/spanned cells, repeat the value in each position it spans so every row has the same number of cells as the header.
- Keep empty cells as empty strings rather than dropping them.
- Distinguish genuine tables from layout columns, code blocks, and figures — extract only real tabular data.
- If a header row is repeated mid-table (page break artifact), keep the data rows but do not duplicate the header.

Report ALL tables by calling the record_tables tool. If there are no tables on this page, call it with an empty list."""


def compute_dpi(page) -> float:
    """Pick a render DPI so the page's long edge lands near TARGET_LONG_EDGE_PX."""
    rect = page.rect
    long_edge_pts = max(rect.width, rect.height)  # points (1/72 inch)
    if long_edge_pts <= 0:
        return 150.0
    dpi = TARGET_LONG_EDGE_PX * 72.0 / long_edge_pts
    return max(MIN_DPI, min(MAX_DPI, dpi))


def pdf_page_to_base64(page) -> str:
    """Render an already-open PDF page to a base64-encoded PNG image."""
    dpi = compute_dpi(page)
    mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img_bytes = pix.tobytes("png")
    return base64.standard_b64encode(img_bytes).decode("utf-8")


def extract_tables_from_page(
    client: anthropic.Anthropic,
    image_b64: str,
    page_num: int,
    pdf_filename: str,
) -> list[dict]:
    """Send one page image to Claude and return the list of structured tables."""
    messages = [
        {
            "role": "user",
            "content": [
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/png",
                        "data": image_b64,
                    },
                },
                {"type": "text", "text": EXTRACTION_PROMPT},
            ],
        }
    ]

    last_err: Optional[Exception] = None
    for attempt in range(MAX_RETRIES):
        try:
            # Streaming avoids request timeouts on dense pages with large tables.
            # thinking=adaptive improves transcription accuracy; tool_choice stays
            # "auto" because forced tool use is incompatible with extended thinking.
            with client.messages.stream(
                model=MODEL,
                max_tokens=16000,
                thinking={"type": "adaptive"},
                tools=[RECORD_TOOL],
                messages=messages,
            ) as stream:
                final = stream.get_final_message()
            return _tables_from_message(final, page_num, pdf_filename)
        except (anthropic.APIStatusError, anthropic.APIConnectionError) as e:
            last_err = e
            wait = 2 ** attempt
            print(f"    [retry {attempt + 1}/{MAX_RETRIES}] {type(e).__name__}; waiting {wait}s")
            time.sleep(wait)

    print(f"    [error] page {page_num + 1} failed after {MAX_RETRIES} attempts: {last_err}")
    return []


def _tables_from_message(message, page_num: int, pdf_filename: str) -> list[dict]:
    """Pull the record_tables tool input out of a finished message."""
    tables: list[dict] = []
    for block in message.content:
        if getattr(block, "type", None) == "tool_use" and block.name == "record_tables":
            tables = list(block.input.get("tables", []))
            break
    else:
        # Fallback: model answered in prose JSON instead of calling the tool.
        text = "".join(
            b.text for b in message.content if getattr(b, "type", None) == "text"
        )
        tables = _salvage_json_tables(text)

    for t in tables:
        t["page_number"] = page_num + 1
        t["source_file"] = pdf_filename
    return tables


def _salvage_json_tables(text: str) -> list[dict]:
    """Best-effort parse of a JSON tables payload embedded in prose."""
    m = re.search(r"```(?:json)?\s*([\s\S]*?)```", text)
    candidate = m.group(1) if m else text
    for start_char, end_char in (("{", "}"), ("[", "]")):
        s = candidate.find(start_char)
        e = candidate.rfind(end_char)
        if s != -1 and e > s:
            try:
                data = json.loads(candidate[s : e + 1])
                if isinstance(data, dict):
                    return list(data.get("tables", []))
                if isinstance(data, list):
                    return data
            except json.JSONDecodeError:
                continue
    return []


def stitch_continuations(tables: list[dict]) -> list[dict]:
    """Merge tables that span page breaks into single tables.

    A table flagged continued_from_previous is appended onto the most recent
    open table when their headers are compatible (identical, or one side empty).
    """
    merged: list[dict] = []
    for t in tables:
        prev = merged[-1] if merged else None
        can_merge = (
            prev is not None
            and t.get("continued_from_previous")
            and prev.get("continues_on_next")
            and _headers_compatible(prev.get("headers", []), t.get("headers", []))
        )
        if can_merge:
            prev["rows"].extend(t.get("rows", []))
            # Carry footnotes forward and update where the table now ends.
            prev["footnotes"] = prev.get("footnotes", []) + t.get("footnotes", [])
            prev["continues_on_next"] = t.get("continues_on_next", False)
            prev.setdefault("spans_pages", [prev.get("page_number")])
            prev["spans_pages"].append(t.get("page_number"))
        else:
            merged.append(t)
    return merged


def _headers_compatible(h1: list, h2: list) -> bool:
    if not h1 or not h2:
        return True
    if len(h1) != len(h2):
        return False
    norm = lambda h: [str(x).strip().lower() for x in h]
    return norm(h1) == norm(h2)


def sanitize_filename(name: str, max_len: int = 60) -> str:
    name = re.sub(r'[\\/*?:"<>|]', "_", str(name))
    name = re.sub(r"\s+", "_", name.strip())
    return name[:max_len] if name else "untitled"


def write_table_csv(table: dict, output_dir: Path, global_index: int) -> str:
    """Write a single table to its own CSV with an Azure-DI-style metadata block."""
    title = table.get("title") or f"table_{global_index}"
    page = table.get("page_number", 0)
    filename = f"p{page:04d}_t{global_index:04d}_{sanitize_filename(title)}.csv"
    filepath = output_dir / filename

    headers = table.get("headers", [])
    rows = table.get("rows", [])

    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["# Source File", table.get("source_file", "")])
        spans = table.get("spans_pages")
        writer.writerow(["# Page(s)", ", ".join(map(str, spans)) if spans else page])
        writer.writerow(["# Global Table Index", global_index])
        writer.writerow(["# Title", title])
        writer.writerow(["# Context Label", table.get("context_label") or ""])
        for i, fn in enumerate(table.get("footnotes", []), 1):
            writer.writerow([f"# Footnote {i}", fn])
        writer.writerow([])
        if headers:
            writer.writerow(headers)
        for row in rows:
            writer.writerow(row)

    return str(filepath)


def write_combined_csv(all_tables: list[dict], output_dir: Path, pdf_stem: str) -> str:
    """Write every table into one long-format CSV for spreadsheet analysis."""
    filepath = output_dir / f"{pdf_stem}__ALL_TABLES.csv"
    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow([
            "global_table_index", "source_file", "page_number", "spans_pages",
            "title", "context_label", "footnotes", "headers",
            "row_index", "row_data",
        ])
        for g_idx, table in enumerate(all_tables, 1):
            spans = table.get("spans_pages")
            base = [
                g_idx,
                table.get("source_file", ""),
                table.get("page_number", ""),
                ", ".join(map(str, spans)) if spans else "",
                table.get("title") or "",
                table.get("context_label") or "",
                " | ".join(table.get("footnotes", [])),
                " | ".join(str(h) for h in table.get("headers", [])),
            ]
            rows = table.get("rows", [])
            if not rows:
                writer.writerow(base + ["", ""])
            for r_idx, row in enumerate(rows, 1):
                writer.writerow(base + [r_idx, " | ".join(str(c) for c in row)])
    return str(filepath)


def write_manifest(all_tables: list[dict], output_dir: Path, pdf_stem: str) -> str:
    """Write the full structured extraction as JSON (Azure-DI-like manifest)."""
    filepath = output_dir / f"{pdf_stem}__manifest.json"
    manifest = {
        "source": pdf_stem,
        "table_count": len(all_tables),
        "tables": [
            {
                "global_table_index": i + 1,
                "page_number": t.get("page_number"),
                "spans_pages": t.get("spans_pages"),
                "title": t.get("title"),
                "context_label": t.get("context_label"),
                "footnotes": t.get("footnotes", []),
                "headers": t.get("headers", []),
                "row_count": len(t.get("rows", [])),
                "column_count": len(t.get("headers", [])) or (
                    len(t["rows"][0]) if t.get("rows") else 0
                ),
                "rows": t.get("rows", []),
            }
            for i, t in enumerate(all_tables)
        ],
    }
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    return str(filepath)


def process_pdf(pdf_path: Path, output_root: Path, client: anthropic.Anthropic) -> int:
    """Extract all tables from one PDF. Returns the number of tables written."""
    pdf_filename = pdf_path.name
    pdf_stem = pdf_path.stem

    print(f"\n{'=' * 60}\nProcessing: {pdf_filename}")
    doc = fitz.open(str(pdf_path))
    total_pages = len(doc)
    print(f"Total pages: {total_pages}")

    out_dir = output_root / sanitize_filename(pdf_stem)
    out_dir.mkdir(parents=True, exist_ok=True)

    page_tables: list[dict] = []
    for page_num in range(total_pages):
        print(f"  Page {page_num + 1}/{total_pages}...", end=" ", flush=True)
        image_b64 = pdf_page_to_base64(doc[page_num])
        tables = extract_tables_from_page(client, image_b64, page_num, pdf_filename)
        print(f"found {len(tables)} table(s)" if tables else "no tables")
        page_tables.extend(tables)
    doc.close()

    # Stitch cross-page tables, then persist.
    all_tables = stitch_continuations(page_tables)

    for g_idx, table in enumerate(all_tables, 1):
        csv_path = write_table_csv(table, out_dir, g_idx)
        print(f"    -> {Path(csv_path).name}")

    if all_tables:
        print(f"\n  Combined : {write_combined_csv(all_tables, out_dir, pdf_stem)}")
        print(f"  Manifest : {write_manifest(all_tables, out_dir, pdf_stem)}")
        print(f"  Tables   : {len(all_tables)} "
              f"(from {len(page_tables)} page-level detections)")
    else:
        print(f"\n  No tables found in {pdf_filename}")

    return len(all_tables)


def resolve_inputs() -> tuple[list[Path], Path]:
    """Get PDF list and output folder from argv or interactive prompts."""
    args = sys.argv[1:]
    pdf_arg = args[0] if len(args) >= 1 else input(
        "\nEnter path to a PDF file or a folder of PDFs: ").strip()
    out_arg = args[1] if len(args) >= 2 else input(
        "Enter output folder for CSV results: ").strip()

    pdf_path = Path(pdf_arg.strip().strip('"').strip("'"))
    if not pdf_path.exists():
        sys.exit(f"Error: path not found: {pdf_path}")

    if pdf_path.is_file():
        if pdf_path.suffix.lower() != ".pdf":
            sys.exit("Error: file must be a .pdf")
        pdf_files = [pdf_path]
    else:
        pdf_files = sorted(
            p for p in pdf_path.iterdir() if p.suffix.lower() == ".pdf"
        )
        if not pdf_files:
            sys.exit(f"Error: no PDF files found in {pdf_path}")
        print(f"Found {len(pdf_files)} PDF file(s)")

    output_dir = Path(out_arg.strip().strip('"').strip("'"))
    output_dir.mkdir(parents=True, exist_ok=True)
    return pdf_files, output_dir


def main():
    print("PDF Table Extractor — powered by Claude Vision")
    print("=" * 60)

    pdf_files, output_dir = resolve_inputs()
    print(f"Output folder: {output_dir.resolve()}")

    api_key = os.environ.get("ANTHROPIC_API_KEY") or input(
        "Enter your Anthropic API key: ").strip()
    client = anthropic.Anthropic(api_key=api_key)

    total = sum(process_pdf(p, output_dir, client) for p in pdf_files)

    print(f"\n{'=' * 60}\nExtraction complete.")
    print(f"Total tables extracted: {total}")
    print(f"Results saved to: {output_dir.resolve()}")


if __name__ == "__main__":
    main()
