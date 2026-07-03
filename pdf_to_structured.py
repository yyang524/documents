#!/usr/bin/env python3
"""
PDF -> structured Markdown + CSV, offline (PyMuPDF only).

Built for long Word-generated PDFs (IMF country reports and similar) with
complex, often borderless tables and many footnotes. Outputs are meant to be
read by an LLM assistant (e.g. GitHub Copilot in VS Code):

  <output>/<pdf-name>/
      <pdf-name>.md          full document text, page-marked, all footnotes
      tables/
          t001_p0039_Table_1_....csv     one CSV per table
      tables_index.csv       caption, pages, status of every table found

How tables are found — caption-driven, not detection-driven:
  A table is anything whose caption line contains "Table N" (e.g. "Table 3.",
  "Text Table 2.", "Annex VIII. Table 1.", "Table 2a."). Captions containing
  "Figure" are charts and are ignored. The table band runs from the caption
  down to the next caption / end of page; footnote lines ("Sources:", "1/ ...",
  "Note:") inside the band are captured as footnotes, not data.

How each band becomes a grid — three strategies, first one that works wins:
  1. ruled     - PyMuPDF find_tables() clipped to the band (bordered tables)
  2. numeric   - right edges of numeric tokens cluster into columns
                 (right-aligned statistical tables with no ruling lines)
  3. gaps      - whitespace-gap profile across the band (text-only tables,
                 e.g. reporting-requirement or structural-benchmark tables)

Tables pasted as pictures (image, no text layer) cannot be read offline
without OCR; they are flagged in the CSV and the index instead of being
silently dropped.

Multi-page tables: a caption on the next page whose base matches the previous
caption (ignoring "(continued)"/"(concluded)") is merged into one CSV.

Dependencies: PyMuPDF only.  python -m pip install --user pymupdf
Usage:        python pdf_to_structured.py [PDF_OR_FOLDER] [OUTPUT_FOLDER]
"""

import csv
import re
import sys
from pathlib import Path
from typing import Optional

try:
    import fitz  # pymupdf
except ImportError:
    sys.exit("PyMuPDF is required:  python -m pip install --user pymupdf")


# --------------------------------------------------------------------------
# Tunables
# --------------------------------------------------------------------------
ROW_TOL = 3.0            # points: words within this vertical distance = one row
NUM_COL_TOL = 8.0        # points: cluster width for numeric right edges
GAP_MIN = 8              # points: minimum whitespace gap = column separator
MIN_NUMS_PER_ROW = 3     # numeric tokens needed for a row to vote on columns
MIN_GRID_ROWS = 2        # fewer data rows than this -> extraction failed
IMAGE_TABLE_MAX_WORDS = 25   # a band with fewer words + an image = picture table

CAPTION_RE = re.compile(
    r"^\s*(?:Annex\s+[IVXLC]+\.?\s*)?(?:Text\s+)?Table\s+\d+[a-z]?\s*[.:]", re.I)
FIGURE_RE = re.compile(r"\bFigure\b", re.I)
CONT_RE = re.compile(r"\(\s*con(?:tinued|cluded)\s*\)", re.I)
FOOTNOTE_START_RE = re.compile(r"^(Sources?\s*:|Notes?\s*:|\d+/\s)")
NUMBER_RE = re.compile(r"^[-(]?[\d,]+\.?\d*\)?%?$")


def is_number(text: str) -> bool:
    t = text.strip().replace("–", "-").replace("−", "-")
    return bool(re.search(r"\d", t)) and bool(NUMBER_RE.match(t))


# --------------------------------------------------------------------------
# Geometry helpers
# --------------------------------------------------------------------------
def group_rows(words: list) -> list:
    """Group word tuples (x0,y0,x1,y1,text,...) into visual rows."""
    rows, cur, cur_y = [], [], None
    for w in sorted(words, key=lambda w: ((w[1] + w[3]) / 2, w[0])):
        yc = (w[1] + w[3]) / 2
        if cur_y is None or abs(yc - cur_y) <= ROW_TOL:
            cur.append(w)
            cur_y = yc if cur_y is None else (cur_y + yc) / 2
        else:
            rows.append(sorted(cur, key=lambda w: w[0]))
            cur, cur_y = [w], yc
    if cur:
        rows.append(sorted(cur, key=lambda w: w[0]))
    return rows


def row_text(row: list) -> str:
    return " ".join(w[4] for w in row)


def cluster_1d(values: list, tol: float) -> list:
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


# --------------------------------------------------------------------------
# Grid strategies
# --------------------------------------------------------------------------
def grid_score(grid: Optional[list]) -> float:
    """Quality of an extracted grid: reward well-separated cells and column
    count, penalize 'jammed' cells holding many numbers (a sign that several
    real columns collapsed into one)."""
    if not grid:
        return -1
    filled = [c for r in grid for c in r if c]
    if not filled:
        return -1
    jammed = sum(1 for c in filled
                 if len(re.findall(r"[-(]?\d[\d,]*\.?\d*\)?%?(?:\s|$)", c)) >= 4)
    n_cols = max(len(r) for r in grid)
    return len(filled) - 4 * jammed + 2 * n_cols


def compact_grid(grid: list) -> list:
    """Clean up column noise: sparse stray columns (created by e.g. a bold
    header whose glyph edges differ slightly from the data edges) are merged
    into the adjacent real column when that doesn't overwrite anything, and
    fully empty columns are dropped."""
    if not grid:
        return grid
    width = max(len(r) for r in grid)
    grid = [r + [""] * (width - len(r)) for r in grid]
    fills = [sum(1 for r in grid if r[c]) for c in range(width)]
    real = max(fills) if fills else 0

    for c in range(width):
        if not (0 < fills[c] <= max(1, int(0.15 * real))):
            continue
        for n in (c + 1, c - 1):                      # prefer right neighbor
            if 0 <= n < width and fills[n] > 0.5 * real and \
                    all(not r[n] for r in grid if r[c]):
                for r in grid:
                    if r[c]:
                        r[n], r[c] = r[c], ""
                fills[n] += fills[c]
                fills[c] = 0
                break

    keep = [c for c in range(width) if any(r[c] for r in grid)]
    return [[r[c] for c in keep] for r in grid]


def grid_from_ruled(page, band) -> Optional[list]:
    """Strategy 1: PyMuPDF's line-based detector, clipped to the band."""
    try:
        finder = page.find_tables(clip=fitz.Rect(band))
    except Exception:
        return None
    best = None
    for t in finder.tables:
        rows = [[("" if c is None else re.sub(r"\s+", " ", str(c)).strip())
                 for c in r] for r in t.extract()]
        rows = [r for r in rows if any(r)]
        if len(rows) >= MIN_GRID_ROWS and max(len(r) for r in rows) >= 2:
            if best is None or len(rows) > len(best):
                best = rows
    return best


def grid_from_numeric_edges(rows: list) -> Optional[list]:
    """Strategy 2: columns = clustered right edges of numeric tokens."""
    edges = [w[2] for r in rows for w in r
             if is_number(w[4]) and sum(is_number(x[4]) for x in r) >= MIN_NUMS_PER_ROW]
    cols = cluster_1d(edges, NUM_COL_TOL)
    if len(cols) < 2:
        return None
    label_right = min(cols) - 30

    out = []
    for r in rows:
        label, cells = [], [""] * len(cols)
        for w in r:
            if is_number(w[4]):
                ci = min(range(len(cols)), key=lambda i: abs(cols[i] - w[2]))
                cells[ci] = (cells[ci] + " " + w[4]).strip()
            elif w[2] < label_right or not any(cells):
                label.append(w[4])
        out.append([" ".join(label).strip()] + cells)
    out = [r for r in out if any(r)]
    data = [r for r in out if sum(1 for c in r[1:] if c) >= 2]
    return out if len(data) >= MIN_GRID_ROWS else None


def grid_from_gaps(rows: list) -> Optional[list]:
    """Strategy 3: whitespace-gap profile (for text-only tables).

    Column separators are x-ranges that stay empty across (nearly) all rows.
    Wide rows spanning the whole band (titles, notes) are left out of the vote.
    """
    body = [r for r in rows if len(r) >= 2]
    if len(body) < MIN_GRID_ROWS:
        return None
    x_lo = min(w[0] for r in body for w in r)
    x_hi = max(w[2] for r in body for w in r)
    width = int(x_hi - x_lo) + 1
    if width <= 0:
        return None

    occupancy = [0] * width
    for r in body:
        for w in r:
            for x in range(max(0, int(w[0] - x_lo)), min(width, int(w[2] - x_lo))):
                occupancy[x] += 1
    # a gap = run of bins used by (almost) no row
    allow = max(1, int(len(body) * 0.08))
    gaps, i = [], 0
    while i < width:
        if occupancy[i] <= allow:
            j = i
            while j < width and occupancy[j] <= allow:
                j += 1
            if j - i >= GAP_MIN and i > 0 and j < width:
                gaps.append((i + x_lo + j + x_lo) / 2)
            i = j
        else:
            i += 1
    if not gaps:
        return None

    bounds = [x_lo - 1] + gaps + [x_hi + 1]
    out = []
    for r in rows:
        cells = [""] * (len(bounds) - 1)
        for w in r:
            xc = (w[0] + w[2]) / 2
            for ci in range(len(bounds) - 1):
                if bounds[ci] <= xc < bounds[ci + 1]:
                    cells[ci] = (cells[ci] + " " + w[4]).strip()
                    break
        out.append(cells)
    out = [r for r in out if any(r)]
    multi = [r for r in out if sum(1 for c in r if c) >= 2]
    return out if len(multi) >= MIN_GRID_ROWS else None


# --------------------------------------------------------------------------
# Caption-driven extraction
# --------------------------------------------------------------------------
def find_captions(blocks: list) -> list:
    """Return [(y_top, y_bottom, caption_text)] for 'Table N' captions."""
    captions = []
    for b in blocks:
        # examine block line by line so a caption inside a text block is found
        y_span = b[3] - b[1]
        lines = [ln for ln in b[4].split("\n") if ln.strip()]
        for k, ln in enumerate(lines):
            ln_clean = re.sub(r"\s+", " ", ln).strip()
            if not CAPTION_RE.match(ln_clean) or FIGURE_RE.search(ln_clean):
                continue
            if len(ln_clean) > 110:      # a sentence mentioning a table, not a title
                continue
            frac = (k + 1) / len(lines)
            y_bot = b[1] + y_span * frac
            captions.append((b[1], y_bot, ln_clean))
            break  # one caption per block is enough
    captions.sort()
    return captions


def looks_like_reference(caption: str) -> bool:
    """True for prose that merely mentions a table ('... in Table 2 of ...')."""
    return not CAPTION_RE.match(caption)


def split_footnotes(rows: list) -> tuple:
    """Split band rows into (table_rows, footnote_lines).

    Footnotes start at the first row matching Sources:/Notes:/n/ and run to
    the end of the band.
    """
    for i, r in enumerate(rows):
        if FOOTNOTE_START_RE.match(row_text(r)):
            notes, cur = [], ""
            for fr in rows[i:]:
                t = row_text(fr)
                if FOOTNOTE_START_RE.match(t):
                    if cur:
                        notes.append(cur)
                    cur = t
                else:
                    cur += " " + t
            if cur:
                notes.append(cur)
            return rows[:i], notes
    return rows, []


def extract_band(page, layout: dict, y_top: float, y_bottom: float) -> dict:
    """Extract one table band into {'grid','footnotes','status','strategy'}."""
    band = (0, y_top, layout["width"], y_bottom)
    words = [w for w in layout["words"]
             if y_top <= (w[1] + w[3]) / 2 <= y_bottom]

    # picture check: almost no words but a large image on the page
    if len(words) <= IMAGE_TABLE_MAX_WORDS:
        for xref, *_ in page.get_images():
            for r in page.get_image_rects(xref):
                if abs(r) > 10000:
                    return {"grid": None, "footnotes": [], "strategy": None,
                            "status": "image_table_needs_ocr"}

    rows = group_rows(words)
    # drop page-number rows and stray caption/continuation-marker lines
    rows = [r for r in rows
            if not re.fullmatch(r"\d{1,3}", row_text(r))
            and not CAPTION_RE.match(row_text(r))]
    rows, footnotes = split_footnotes(rows)
    if not rows:
        return {"grid": None, "footnotes": footnotes, "strategy": None,
                "status": "no_text_rows"}

    # Try all strategies and keep the best-scoring grid (a strategy can
    # "succeed" with a poor grid, e.g. ruled detection collapsing a
    # borderless table into one jammed column).
    candidates = [("numeric", grid_from_numeric_edges(rows)),
                  ("gaps", grid_from_gaps(rows))]
    if not layout["rotated"]:   # ruled detection needs original coordinates
        candidates.insert(0, ("ruled", grid_from_ruled(page, band)))
    best_name, best_grid, best_score = None, None, -1
    for name, grid in candidates:
        s = grid_score(grid)
        if s > best_score:
            best_name, best_grid, best_score = name, grid, s
    if best_grid:
        return {"grid": compact_grid(best_grid), "footnotes": footnotes,
                "strategy": best_name, "status": "ok"}
    # last resort: keep the raw lines, one column, so nothing is lost
    return {"grid": [[row_text(r)] for r in rows], "footnotes": footnotes,
            "strategy": "lines", "status": "unstructured_fallback"}


def caption_base(caption: str) -> str:
    base = CONT_RE.sub("", caption)
    base = re.sub(r"\d+\s*/?\s*$", "", base)      # trailing footnote marks
    return re.sub(r"\s+", " ", base).strip().lower()


BOTTOM_MARKER_FRAC = 0.80   # captions in the bottom 20% = "continues next page"


def page_layout(page) -> dict:
    """Words + text blocks in reading coordinates.

    Landscape tables in these PDFs are stored as vertical text on a portrait
    page (set_rotation does not change extraction coordinates in PyMuPDF), so
    when vertical text dominates we rotate all coordinates ourselves.
    Returns {'words', 'blocks', 'width', 'height', 'rotated'} where each word
    is (x0,y0,x1,y1,text) and each block is (x0,y0,x1,y1,text).
    """
    d = page.get_text("dict")
    up = down = horiz = 0
    for b in d.get("blocks", []):
        for ln in b.get("lines", []):
            dx, dy = ln["dir"]
            if abs(dy) > abs(dx):
                up, down = (up, down + 1) if dy > 0 else (up + 1, down)
            else:
                horiz += 1
    words = [w[:5] for w in page.get_text("words") if w[4].strip()]
    blocks = [(b[0], b[1], b[2], b[3], b[4] or "")
              for b in page.get_text("blocks")
              if (b[6] if len(b) > 6 else 0) == 0]
    W, H = page.rect.width, page.rect.height
    vert = up + down
    if vert <= max(3, 2 * horiz):
        return {"words": words, "blocks": blocks,
                "width": W, "height": H, "rotated": False}

    if up >= down:   # text reads bottom-up: rotate 90 deg clockwise
        tr = lambda x0, y0, x1, y1: (H - y1, x0, H - y0, x1)
    else:            # text reads top-down: rotate 90 deg counter-clockwise
        tr = lambda x0, y0, x1, y1: (y0, W - x1, y1, W - x0)
    words = [(*tr(w[0], w[1], w[2], w[3]), w[4]) for w in words]
    blocks = [(*tr(b[0], b[1], b[2], b[3]), b[4]) for b in blocks]
    return {"words": words, "blocks": blocks,
            "width": H, "height": W, "rotated": True}


def extract_pdf_tables(doc, pdf_name: str) -> list:
    """Walk the document, return list of table dicts (multi-page merged)."""
    tables: list = []
    for pno in range(len(doc)):
        page = doc[pno]
        layout = page_layout(page)
        all_caps = find_captions(layout["blocks"])
        # A caption near the page bottom is a running "…(continued)" marker
        # announcing that the table resumes on the next page — not a new band.
        cutoff = layout["height"] * BOTTOM_MARKER_FRAC
        caps = [c for c in all_caps if c[0] < cutoff]
        has_bottom_marker = any(c[0] >= cutoff for c in all_caps)

        prev = tables[-1] if tables else None
        if not caps:
            # Caption-less page: belongs to the previous table only if that
            # table was left open by a bottom "(continued)" marker.
            if (prev is not None and prev.get("open")
                    and pno + 1 == prev["pages"][-1] + 1):
                band = extract_band(page, layout, 0, layout["height"])
                if band["grid"]:
                    prev["grid"] = (prev["grid"] or []) + band["grid"]
                    prev["footnotes"].extend(f for f in band["footnotes"]
                                             if f not in prev["footnotes"])
                    prev["pages"].append(pno + 1)
                    prev["open"] = has_bottom_marker
            continue

        for i, (cy0, cy1, caption) in enumerate(caps):
            y_end = caps[i + 1][0] - 2 if i + 1 < len(caps) else layout["height"]
            band = extract_band(page, layout, cy1, y_end)
            entry = {
                "caption": caption,
                "pages": [pno + 1],
                "grid": band["grid"],
                "footnotes": band["footnotes"],
                "strategy": band["strategy"],
                "status": band["status"],
                "source_file": pdf_name,
                # only the last band on the page can run onto the next one
                "open": has_bottom_marker and i == len(caps) - 1,
            }
            # merge "(continued)/(concluded)" — same caption base, next page
            prev = tables[-1] if tables else None
            if (prev is not None
                    and pno + 1 - prev["pages"][-1] in (0, 1)
                    and pno + 1 not in prev["pages"]
                    and caption_base(caption) == caption_base(prev["caption"])
                    and prev["grid"] and entry["grid"]):
                prev["grid"].extend(entry["grid"])
                prev["footnotes"].extend(f for f in entry["footnotes"]
                                         if f not in prev["footnotes"])
                prev["pages"].append(pno + 1)
                prev["open"] = entry["open"]
            else:
                tables.append(entry)
    for t in tables:
        t.pop("open", None)
    return tables


# --------------------------------------------------------------------------
# Full-text Markdown (with all footnotes)
# --------------------------------------------------------------------------
def write_document_markdown(doc, out_path: Path, pdf_name: str) -> None:
    parts = [f"# {pdf_name}\n"]
    for pno in range(len(doc)):
        text = doc[pno].get_text("text").rstrip()
        parts.append(f"\n\n---\n\n## Page {pno + 1}\n")
        parts.append(text if text else "*(no extractable text on this page)*")
    out_path.write_text("\n".join(parts), encoding="utf-8")


# --------------------------------------------------------------------------
# CSV output
# --------------------------------------------------------------------------
def sanitize(name: str, max_len: int = 50) -> str:
    name = re.sub(r"[^\w\s\-.]", "", name).strip()
    return re.sub(r"\s+", "_", name)[:max_len].rstrip("_.") or "table"


def write_table_csv(table: dict, out_dir: Path, index: int) -> Path:
    fn = f"t{index:03d}_p{table['pages'][0]:04d}_{sanitize(table['caption'])}.csv"
    path = out_dir / fn
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["# Caption", table["caption"]])
        w.writerow(["# Source", table["source_file"]])
        w.writerow(["# Page(s)", ", ".join(map(str, table["pages"]))])
        w.writerow(["# Extraction", table["strategy"] or table["status"]])
        for i, fn_text in enumerate(table["footnotes"], 1):
            w.writerow([f"# Footnote {i}", fn_text])
        w.writerow([])
        if table["status"] == "image_table_needs_ocr":
            w.writerow(["[This table is a pasted picture with no text layer. "
                        "Offline extraction is impossible without OCR.]"])
        elif not table["grid"]:
            w.writerow([f"[No extractable rows found ({table['status']}).]"])
        else:
            width = max(len(r) for r in table["grid"])
            for row in table["grid"]:
                w.writerow(row + [""] * (width - len(row)))
    return path


def write_index(tables: list, out_dir: Path) -> Path:
    path = out_dir / "tables_index.csv"
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["index", "caption", "pages", "status", "strategy",
                    "rows", "footnotes"])
        for i, t in enumerate(tables, 1):
            w.writerow([i, t["caption"], ", ".join(map(str, t["pages"])),
                        t["status"], t["strategy"] or "",
                        len(t["grid"]) if t["grid"] else 0, len(t["footnotes"])])
    return path


# --------------------------------------------------------------------------
# Driver
# --------------------------------------------------------------------------
def process(pdf_path: Path, out_root: Path) -> None:
    doc = fitz.open(str(pdf_path))
    out_dir = out_root / pdf_path.stem
    tbl_dir = out_dir / "tables"
    tbl_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n=== {pdf_path.name} ({len(doc)} pages) ===")
    md_path = out_dir / f"{pdf_path.stem}.md"
    write_document_markdown(doc, md_path, pdf_path.name)
    print(f"  text  -> {md_path}")

    tables = extract_pdf_tables(doc, pdf_path.name)
    for i, t in enumerate(tables, 1):
        p = write_table_csv(t, tbl_dir, i)
        flag = "" if t["status"] == "ok" else f"   [{t['status']}]"
        print(f"  table -> {p.name}{flag}")
    write_index(tables, out_dir)

    ok = sum(1 for t in tables if t["status"] == "ok")
    img = sum(1 for t in tables if t["status"] == "image_table_needs_ocr")
    print(f"  {len(tables)} tables: {ok} extracted, {img} image-based (need OCR), "
          f"{len(tables) - ok - img} other")
    doc.close()


def main() -> None:
    args = sys.argv[1:]
    src = Path(args[0]) if args else Path(input("PDF file or folder: ").strip())
    out = Path(args[1]) if len(args) > 1 else Path(input("Output folder: ").strip())
    if src.is_file():
        pdfs = [src]
    else:
        pdfs = sorted(src.glob("*.pdf")) + sorted(src.glob("*.PDF"))
    if not pdfs:
        sys.exit("No PDF found.")
    out.mkdir(parents=True, exist_ok=True)
    for p in pdfs:
        process(p, out)


if __name__ == "__main__":
    main()
