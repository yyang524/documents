#!/usr/bin/env python3
"""
PDF Table Extractor — fully offline, no API key required.

Extracts ALL text-searchable tables from PDF documents (including very long
ones) using PyMuPDF's built-in table finder, captures each table's title and
footnotes from the surrounding text, stitches tables that span page breaks,
and exports everything to CSV — Azure Document Intelligence style.

For each table it captures:
  - title / caption (text directly above the table)
  - footnotes / source lines (text directly below the table)
  - column headers and all data rows (merged cells preserved)
  - page number, table index, and cross-page continuation links

Outputs, into a user-specified folder (one sub-folder per PDF):
  - one CSV per high-confidence table (with a metadata header block)
  - one combined CSV per PDF (every high-confidence table, long format)
  - one structured JSON manifest per PDF (full Azure-DI-like dump, including
    low-confidence detections)
  - a low_confidence_review/ sub-folder holding detections that look like
    charts/figures rather than data tables — written, not discarded, so you
    can review them. Nothing is ever dropped.

Note: on chart-heavy PDFs the geometric detector can mistake a chart's axis
labels/legend for a table. The quality gate routes those to the review folder
while keeping large, genuine (even sparse) tables in the main output.

Dependencies:
  - PyMuPDF only. Install without admin rights:
        python -m pip install --user pymupdf
  - No network, no API key, no OCR. Runs entirely offline.
  - Works on text-searchable PDFs (tables made of real, selectable text).

Usage:
    python extract_pdf_tables.py [PDF_PATH_OR_FOLDER] [OUTPUT_FOLDER]
If the two paths are omitted, the script prompts for them interactively.
"""

import sys
import json
import csv
import re
from pathlib import Path
from typing import Optional

try:
    import fitz  # pymupdf
except ImportError:
    sys.exit(
        "PyMuPDF is required. Install it (no admin needed) with:\n"
        "    python -m pip install --user pymupdf"
    )

if not hasattr(fitz.Page, "find_tables"):
    sys.exit(
        "Your PyMuPDF is too old for table detection. Upgrade with:\n"
        "    python -m pip install --user --upgrade pymupdf"
    )


# How close (in points) surrounding text must be to count as a title/footnote.
TITLE_MAX_GAP = 80
FOOTNOTE_MAX_GAP = 70
# Fraction of page height used to flag a table as touching the top/bottom edge
# (i.e. a likely page-spanning continuation).
EDGE_FRACTION = 0.14

# Quality gate: chart-heavy PDFs make PyMuPDF report charts (their axis labels
# and legends) as tiny/empty "tables". The gate is deliberately CONSERVATIVE —
# it only flags detections that carry almost no tabular content, so large but
# sparse real tables (wide assessment matrices, etc.) are never dropped.
# Flagged detections are still written, into a low_confidence_review/ subfolder
# and the JSON manifest, so nothing is lost. Loosen by lowering MIN_NONEMPTY_CELLS.
MIN_COLUMNS = 2
MIN_DATA_ROWS = 1
MIN_NONEMPTY_CELLS = 8      # total non-empty cells across header + data rows


# ---------------------------------------------------------------------------
# Table extraction (PyMuPDF, offline)
# ---------------------------------------------------------------------------
def _clean_cell(value) -> str:
    """Normalize a single cell value."""
    if value is None:
        return ""
    text = str(value).replace("\n", " ").replace("\r", " ")
    return re.sub(r"\s+", " ", text).strip()


def _normalize_rows(rows: list) -> list:
    """Clean all cells and pad every row to the same width."""
    cleaned = [[_clean_cell(c) for c in row] for row in rows]
    if not cleaned:
        return cleaned
    width = max(len(r) for r in cleaned)
    return [r + [""] * (width - len(r)) for r in cleaned]


def _split_header(table) -> tuple[list, list]:
    """Return (headers, data_rows) for a PyMuPDF table object."""
    extracted = _normalize_rows(table.extract())
    if not extracted:
        return [], []

    header = getattr(table, "header", None)
    names = getattr(header, "names", None) if header else None
    external = getattr(header, "external", False) if header else False

    if external and names and any(n for n in names):
        headers = _normalize_rows([list(names)])[0]
        rows = extracted
    else:
        headers = extracted[0]
        rows = extracted[1:]

    # Pad headers to match the widest data row.
    width = max([len(headers)] + [len(r) for r in rows]) if rows else len(headers)
    headers = headers + [""] * (width - len(headers))
    rows = [r + [""] * (width - len(r)) for r in rows]
    return headers, rows


def _is_low_confidence(headers: list, rows: list) -> bool:
    """Heuristic: True only if a detection carries almost no tabular content
    (the signature of a chart/figure picked up by the geometric detector).
    Large-but-sparse real tables are intentionally NOT flagged."""
    grid = ([headers] if headers else []) + rows
    if not grid:
        return True
    n_cols = max(len(r) for r in grid)
    n_filled = sum(1 for r in grid for c in r if c)
    if n_cols < MIN_COLUMNS or len(rows) < MIN_DATA_ROWS:
        return True
    if n_filled < MIN_NONEMPTY_CELLS:
        return True
    return False


def _is_chart_label(text: str) -> bool:
    """True for stray chart axis/legend text that should not be a title/footnote."""
    t = text.strip()
    if re.fullmatch(r"\d{4}", t):                 # a bare year, e.g. 2026
        return True
    if re.fullmatch(r"[-+]?\d+(?:[.,]\d+)?%?", t):  # a bare number/percent
        return True
    if re.fullmatch(r"[A-Z]{2,4}", t):            # a country/series code, e.g. ITA
        return True
    return False


def _clean_title(title: Optional[str]) -> Optional[str]:
    """Drop running-header noise (e.g. a lone all-caps word like 'GERMANY')."""
    if not title:
        return None
    if re.fullmatch(r"[A-Z][A-Z .,&'-]{1,40}", title) and " " not in title.strip():
        return None
    return title


def _context_for_table(page, bbox) -> tuple[Optional[str], list]:
    """Find the title (text just above) and footnotes (text just below) a table."""
    tx0, ty0, tx1, ty1 = bbox

    text_blocks = []
    for b in page.get_text("blocks"):
        # block tuple: (x0, y0, x1, y1, text, block_no, block_type)
        block_type = b[6] if len(b) > 6 else 0
        if block_type != 0:
            continue
        text = (b[4] or "").strip()
        if text:
            text_blocks.append((b[0], b[1], b[2], b[3], text))

    def overlaps_x(x0: float, x1: float) -> bool:
        return x0 < tx1 and x1 > tx0

    # Title: the closest text block ending just above the table top.
    title = None
    best_gap = None
    for x0, y0, x1, y1, text in text_blocks:
        if y1 <= ty0 + 2 and overlaps_x(x0, x1):
            gap = ty0 - y1
            if 0 <= gap < TITLE_MAX_GAP and (best_gap is None or gap < best_gap):
                best_gap = gap
                title = re.sub(r"\s+", " ", text).strip()

    # Footnotes: text blocks starting just below the table bottom.
    below = sorted(
        (y0, text)
        for x0, y0, x1, y1, text in text_blocks
        if y0 >= ty1 - 2 and overlaps_x(x0, x1) and 0 <= (y0 - ty1) < FOOTNOTE_MAX_GAP
    )
    footnotes = []
    for _, text in below[:4]:
        for line in text.split("\n"):
            line = line.strip()
            if line and not _is_chart_label(line):
                footnotes.append(line)

    return _clean_title(title), footnotes


def extract_tables_from_page(page, page_num: int, pdf_filename: str) -> list[dict]:
    """Extract every table on a page as a structured dict."""
    finder = page.find_tables()
    page_height = page.rect.height
    results = []

    for idx, table in enumerate(finder.tables, 1):
        headers, rows = _split_header(table)
        if not headers and not rows:
            continue

        title, footnotes = _context_for_table(page, table.bbox)
        ty0, ty1 = table.bbox[1], table.bbox[3]

        results.append({
            "table_index": idx,
            "title": title,
            "context_label": None,
            "headers": headers,
            "rows": rows,
            "footnotes": footnotes,
            # Flag tables touching the page edges as possible continuations.
            "continued_from_previous": ty0 < page_height * EDGE_FRACTION,
            "continues_on_next": ty1 > page_height * (1 - EDGE_FRACTION),
            "page_number": page_num + 1,
            "source_file": pdf_filename,
            "low_confidence": _is_low_confidence(headers, rows),
        })

    return results


# ---------------------------------------------------------------------------
# Cross-page stitching
# ---------------------------------------------------------------------------
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
            and prev.get("page_number") != t.get("page_number")
            and _headers_compatible(prev.get("headers", []), t.get("headers", []))
        )
        if can_merge:
            prev["rows"].extend(t.get("rows", []))
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


# ---------------------------------------------------------------------------
# Output writers
# ---------------------------------------------------------------------------
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


def _manifest_entry(t: dict, index: int) -> dict:
    return {
        "global_table_index": index,
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


def write_manifest(all_tables: list[dict], output_dir: Path, pdf_stem: str,
                   low_confidence: Optional[list[dict]] = None) -> str:
    """Write the full structured extraction as JSON (Azure-DI-like manifest)."""
    filepath = output_dir / f"{pdf_stem}__manifest.json"
    manifest = {
        "source": pdf_stem,
        "table_count": len(all_tables),
        "tables": [_manifest_entry(t, i + 1) for i, t in enumerate(all_tables)],
        "low_confidence_count": len(low_confidence or []),
        "low_confidence_tables": [
            _manifest_entry(t, i + 1) for i, t in enumerate(low_confidence or [])
        ],
    }
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    return str(filepath)


# ---------------------------------------------------------------------------
# Driver
# ---------------------------------------------------------------------------
def process_pdf(pdf_path: Path, output_root: Path) -> int:
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
        tables = extract_tables_from_page(doc[page_num], page_num, pdf_filename)
        print(f"found {len(tables)} table(s)" if tables else "no tables")
        page_tables.extend(tables)
    doc.close()

    # Separate genuine data tables from likely chart/figure noise.
    real = [t for t in page_tables if not t.get("low_confidence")]
    low = [t for t in page_tables if t.get("low_confidence")]

    all_tables = stitch_continuations(real)

    for g_idx, table in enumerate(all_tables, 1):
        csv_path = write_table_csv(table, out_dir, g_idx)
        print(f"    -> {Path(csv_path).name}")

    # Low-confidence detections are still written, just kept separate for review.
    if low:
        review_dir = out_dir / "low_confidence_review"
        review_dir.mkdir(parents=True, exist_ok=True)
        for g_idx, table in enumerate(low, 1):
            write_table_csv(table, review_dir, g_idx)

    if all_tables:
        print(f"\n  Combined : {write_combined_csv(all_tables, out_dir, pdf_stem)}")
        print(f"  Manifest : {write_manifest(all_tables, out_dir, pdf_stem, low)}")
        print(f"  Tables   : {len(all_tables)} data table(s) written "
              f"(from {len(real)} high-confidence detections)")
        if low:
            print(f"  Review   : {len(low)} low-confidence detection(s) "
                  f"(likely charts/figures) in low_confidence_review/")
    else:
        print(f"\n  No high-confidence data tables found in {pdf_filename}"
              + (f"; {len(low)} low-confidence detections in low_confidence_review/"
                 if low else ""))

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
    print("PDF Table Extractor — offline (PyMuPDF, no API key)")
    print("=" * 60)

    pdf_files, output_dir = resolve_inputs()
    print(f"Output folder: {output_dir.resolve()}")

    total = sum(process_pdf(p, output_dir) for p in pdf_files)

    print(f"\n{'=' * 60}\nExtraction complete.")
    print(f"Total tables extracted: {total}")
    print(f"Results saved to: {output_dir.resolve()}")


if __name__ == "__main__":
    main()
