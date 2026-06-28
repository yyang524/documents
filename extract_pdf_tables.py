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
  - markdown_cards/: one self-contained Markdown "card" per table (YAML
    front-matter + pipe table + footnotes). This is the LLM-ingestion format:
    each card is bounded and self-describing, so it survives chunked retrieval
    and lines up across documents via a stable table_id.
  - one combined __ALL_TABLES.md per PDF (all cards in one file) — a single
    knowledge source you can drop into Copilot Studio / a RAG index.
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


# ---------------------------------------------------------------------------
# Borderless-table reconstruction (no ruling lines, e.g. IMF statistical
# appendix tables). PyMuPDF's find_tables() keys off vector ruling lines and
# returns nothing on these, so we rebuild the grid from the PDF's own word
# coordinates: numeric columns in such tables are right-aligned, so the right
# edges of number tokens cluster into vertical bands = the columns. Still pure
# PyMuPDF, still offline, no OCR, no new dependencies.
# ---------------------------------------------------------------------------
# Minimum numeric tokens in a row before it helps define column positions, and
# the minimum number of genuine data rows before a reconstruction is accepted.
BORDERLESS_MIN_NUMS_PER_ROW = 3
BORDERLESS_MIN_DATA_ROWS = 4
BORDERLESS_COL_TOL = 8       # points; merge numeric right-edges within this
BORDERLESS_ROW_TOL = 3       # points; merge words into one row within this
BORDERLESS_LABEL_PAD = 30    # points left of the first numeric column = labels

_NUMBER_RE = re.compile(r"^[-(]?[\d,]+\.?\d*\)?%?$")


def _is_number_token(text: str) -> bool:
    """True for a numeric data token like -0.7, 1,234.5, (3.2), 45%, 2026."""
    t = text.strip().replace("–", "-").replace("−", "-")
    return bool(re.search(r"\d", t)) and bool(_NUMBER_RE.match(t))


def _cluster_1d(values: list, tol: float) -> list:
    """Cluster sorted 1-D values into groups within `tol`; return group means."""
    if not values:
        return []
    values = sorted(values)
    groups = [[values[0]]]
    for v in values[1:]:
        if v - groups[-1][-1] <= tol:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [sum(g) / len(g) for g in groups]


def _group_words_into_rows(words: list) -> list:
    """Group word tuples into visual rows by vertical centre."""
    rows, cur, cur_y = [], [], None
    for w in sorted(words, key=lambda w: ((w[1] + w[3]) / 2, w[0])):
        yc = (w[1] + w[3]) / 2
        if cur_y is None or abs(yc - cur_y) <= BORDERLESS_ROW_TOL:
            cur.append(w)
            cur_y = yc if cur_y is None else (cur_y + yc) / 2
        else:
            rows.append(cur)
            cur, cur_y = [w], yc
    if cur:
        rows.append(cur)
    return rows


_CAPTION_RE = re.compile(r"\b(?:Text\s+)?Table\s+\d+\b|\bFigure\s+\d+\b", re.I)


def _find_caption_above(page, table_top: float, max_gap: float = 200) -> Optional[str]:
    """Scan text lines above a borderless table for a 'Table N. …' caption."""
    best = None
    best_y = None
    for b in page.get_text("blocks"):
        if (b[6] if len(b) > 6 else 0) != 0:
            continue
        for line in (b[4] or "").split("\n"):
            line = re.sub(r"\s+", " ", line).strip()
            if not line or not _CAPTION_RE.search(line):
                continue
            y = b[1]
            if y < table_top and (table_top - y) < max_gap:
                # Keep the caption closest to (just above) the table.
                if best_y is None or y > best_y:
                    best, best_y = line, y
    return best


def extract_borderless_tables(page, page_num: int, pdf_filename: str) -> list[dict]:
    """Reconstruct a borderless table from word coordinates (one per page).

    Returns a single-element list (or empty) — these appendix pages hold one
    big table each. Prose pages produce no aligned numeric grid and return [].
    """
    words = [w for w in page.get_text("words") if w[4].strip()]
    if not words:
        return []

    rows = _group_words_into_rows(words)

    # Column positions from right-aligned numeric tokens in number-rich rows.
    edges = [
        w[2] for r in rows
        for w in r
        if _is_number_token(w[4])
        and sum(_is_number_token(x[4]) for x in r) >= BORDERLESS_MIN_NUMS_PER_ROW
    ]
    cols = _cluster_1d(edges, BORDERLESS_COL_TOL)
    if len(cols) < MIN_COLUMNS:
        return []
    first_col_left = min(cols) - BORDERLESS_LABEL_PAD

    def assign(row) -> tuple:
        label_parts, cells = [], [""] * len(cols)
        for w in sorted(row, key=lambda w: w[0]):
            if _is_number_token(w[4]):
                ci = min(range(len(cols)), key=lambda i: abs(cols[i] - w[2]))
                cells[ci] = (cells[ci] + " " + w[4]).strip()
            elif w[2] < first_col_left:
                label_parts.append(w[4])
        return " ".join(label_parts).strip(), cells

    # Build (label, cells) for every row; a "data row" has >=2 filled cells.
    built = [assign(r) for r in rows]
    data_idx = [i for i, (_, cells) in enumerate(built)
                if sum(1 for c in cells if c) >= 2]
    if len(data_idx) < BORDERLESS_MIN_DATA_ROWS:
        return []

    # Restrict to the contiguous table band (first..last data row); rows in
    # between with only a label are kept as section headers.
    band = built[data_idx[0]: data_idx[-1] + 1]
    band = [(lbl, cells) for lbl, cells in band if lbl or any(cells)]

    # Header = leading band rows that are year/number-only (empty label).
    header_cells = []
    body = band
    while body and not body[0][0] and any(body[0][1]):
        header_cells.append(body[0][1])
        body = body[1:]
    if header_cells:
        headers = [""] + [
            " ".join(hc[c] for hc in header_cells if hc[c]).strip()
            for c in range(len(cols))
        ]
    else:
        headers = [""] + [f"col{c + 1}" for c in range(len(cols))]

    table_rows = [[lbl] + cells for lbl, cells in body]
    if not table_rows:
        return []

    # Bounding box of the reconstructed table, for title/footnote context.
    band_words = [w for r in rows for w in r]
    bx0 = min(w[0] for w in band_words)
    bx1 = max(w[2] for w in band_words)
    by0 = min((w[1] for r_i in data_idx for w in rows[r_i]), default=0)
    by1 = max((w[3] for r_i in data_idx for w in rows[r_i]), default=0)
    title, footnotes = _context_for_table(page, (bx0, by0, bx1, by1))
    # The data band starts below the caption, so the nearest line above is a
    # sub-header ("Est.", units). Prefer a real caption line if one sits above.
    caption = _find_caption_above(page, by0)
    if caption:
        title = caption

    page_height = page.rect.height
    return [{
        "table_index": 1,
        "title": title,
        "context_label": None,
        "headers": headers,
        "rows": table_rows,
        "footnotes": footnotes,
        "continued_from_previous": by0 < page_height * EDGE_FRACTION,
        "continues_on_next": by1 > page_height * (1 - EDGE_FRACTION),
        "page_number": page_num + 1,
        "source_file": pdf_filename,
        "low_confidence": _is_low_confidence(headers, table_rows),
        "detection": "borderless",
    }]


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
            "detection": "ruled",
        })

    # Borderless fallback: if the line-based detector found nothing, the page
    # may still hold a borderless table (e.g. IMF statistical appendix).
    if not results:
        results = extract_borderless_tables(page, page_num, pdf_filename)

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
        headers_ok = _headers_compatible(prev.get("headers", []) if prev else [],
                                         t.get("headers", []))
        consecutive = bool(prev) and t.get("page_number") == _last_page(prev) + 1

        # Two ways to recognise a continuation:
        #   (a) geometric — the previous table runs off the bottom and this one
        #       starts at the top of the next page (works for ruled tables);
        #   (b) caption — both pages carry the same caption base, with the
        #       continuation marked "(continued)/(concluded)" (works for the
        #       borderless statistical tables, whose title pushes the data band
        #       away from the page edge so the geometric flags miss them).
        edge_continuation = (
            bool(prev)
            and t.get("continued_from_previous")
            and prev.get("continues_on_next")
            and prev.get("page_number") != t.get("page_number")
        )
        caption_continuation = (
            consecutive
            and _title_base(prev.get("title")) != ""
            and _title_base(prev.get("title")) == _title_base(t.get("title"))
        )
        can_merge = headers_ok and (edge_continuation or caption_continuation)

        if can_merge:
            prev["rows"].extend(t.get("rows", []))
            prev["footnotes"] = prev.get("footnotes", []) + t.get("footnotes", [])
            prev["continues_on_next"] = t.get("continues_on_next", False)
            prev.setdefault("spans_pages", [prev.get("page_number")])
            prev["spans_pages"].append(t.get("page_number"))
        else:
            merged.append(t)
    return merged


def _last_page(table: dict) -> int:
    """Highest page a (possibly already-stitched) table covers."""
    spans = table.get("spans_pages")
    return spans[-1] if spans else table.get("page_number", 0)


def _title_base(title: Optional[str]) -> str:
    """Normalised caption for matching continuations: drop (continued)/
    (concluded) markers and trailing footnote superscripts."""
    if not title:
        return ""
    base = re.sub(r"\((?:con(?:tinued|cluded))\)", "", title, flags=re.I)
    base = re.sub(r"\d+\s*$", "", base)           # trailing footnote digit(s)
    return re.sub(r"\s+", " ", base).strip().lower()


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


def _slugify(text: str) -> str:
    """Stable id from a title, for matching the same table across two documents."""
    text = re.sub(r"\s+", " ", (text or "")).strip().lower()
    text = re.sub(r"[^a-z0-9]+", "-", text).strip("-")
    return text[:80] or "table"


def _units_from_title(title: Optional[str]) -> str:
    """Pull a units hint out of a title, e.g. '(Percent of GDP)'."""
    if not title:
        return ""
    m = re.search(r"\(([^)]*(?:percent|%|gdp|usd|eur|index|growth|ratio|"
                  r"thousands|millions|billions)[^)]*)\)", title, re.IGNORECASE)
    return m.group(1).strip() if m else ""


def _escape_md(cell: str) -> str:
    return str(cell).replace("|", "\\|").strip()


def _table_as_markdown(table: dict) -> str:
    """Render just the pipe-table body (headers + rows) as Markdown."""
    headers = table.get("headers", [])
    rows = table.get("rows", [])
    width = max([len(headers)] + [len(r) for r in rows]) if (headers or rows) else 0
    if width == 0:
        return ""
    head = headers + [""] * (width - len(headers))
    lines = [
        "| " + " | ".join(_escape_md(c) for c in head) + " |",
        "| " + " | ".join(["---"] * width) + " |",
    ]
    for r in rows:
        r = list(r) + [""] * (width - len(r))
        lines.append("| " + " | ".join(_escape_md(c) for c in r) + " |")
    return "\n".join(lines)


def _table_card_markdown(table: dict, global_index: int) -> str:
    """One self-contained Markdown 'card': YAML front-matter + table + footnotes.

    This is the LLM-ingestion unit for chunked, cross-document comparison: each
    card is bounded and self-describing, so it survives retrieval out of order.
    """
    title = table.get("title") or f"Table {global_index}"
    spans = table.get("spans_pages")
    page_str = ", ".join(map(str, spans)) if spans else str(table.get("page_number", ""))
    footnotes = table.get("footnotes", [])

    fm = ["---"]
    fm.append(f"table_id: {_slugify(title)}")
    fm.append(f"source_file: {table.get('source_file', '')}")
    fm.append(f"page: {page_str}")
    fm.append(f"global_table_index: {global_index}")
    fm.append(f'title: "{title.replace(chr(34), chr(39))}"')
    units = _units_from_title(title)
    if units:
        fm.append(f'units: "{units}"')
    fm.append(f"columns: {len(table.get('headers', []))}")
    fm.append(f"rows: {len(table.get('rows', []))}")
    if footnotes:
        fm.append("footnotes:")
        for fn in footnotes:
            fm.append(f'  - "{fn.replace(chr(34), chr(39))}"')
    fm.append("---")

    body = [f"\n# {title}\n", _table_as_markdown(table)]
    if footnotes:
        body.append("\n**Footnotes:**")
        body.extend(f"- {fn}" for fn in footnotes)
    return "\n".join(fm) + "\n" + "\n".join(body) + "\n"


def write_table_markdown(table: dict, output_dir: Path, global_index: int) -> str:
    """Write a single table as a self-contained Markdown card (LLM-ready)."""
    title = table.get("title") or f"table_{global_index}"
    page = table.get("page_number", 0)
    filename = f"p{page:04d}_t{global_index:04d}_{sanitize_filename(title)}.md"
    filepath = output_dir / filename
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(_table_card_markdown(table, global_index))
    return str(filepath)


def write_combined_markdown(all_tables: list[dict], output_dir: Path, pdf_stem: str) -> str:
    """Write all tables into one Markdown file — a single Copilot knowledge source."""
    filepath = output_dir / f"{pdf_stem}__ALL_TABLES.md"
    parts = [f"# Tables extracted from {pdf_stem}\n",
             f"_{len(all_tables)} table(s)._\n"]
    for g_idx, table in enumerate(all_tables, 1):
        parts.append("\n---\n")
        parts.append(_table_card_markdown(table, g_idx))
    with open(filepath, "w", encoding="utf-8") as f:
        f.write("\n".join(parts))
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

    # Markdown "cards" are the LLM-ingestion format (see write_combined_markdown);
    # kept in their own subfolder so the CSV directory stays human/Excel-focused.
    cards_dir = out_dir / "markdown_cards"
    if all_tables:
        cards_dir.mkdir(parents=True, exist_ok=True)

    for g_idx, table in enumerate(all_tables, 1):
        csv_path = write_table_csv(table, out_dir, g_idx)
        write_table_markdown(table, cards_dir, g_idx)
        print(f"    -> {Path(csv_path).name}")

    # Low-confidence detections are still written, just kept separate for review.
    if low:
        review_dir = out_dir / "low_confidence_review"
        review_dir.mkdir(parents=True, exist_ok=True)
        for g_idx, table in enumerate(low, 1):
            write_table_csv(table, review_dir, g_idx)

    if all_tables:
        print(f"\n  Combined : {write_combined_csv(all_tables, out_dir, pdf_stem)}")
        print(f"  Markdown : {write_combined_markdown(all_tables, out_dir, pdf_stem)}")
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
