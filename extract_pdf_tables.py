#!/usr/bin/env python3
"""
PDF Table Extractor using Claude Vision API
Extracts all tables (including image-embedded) from PDF documents with titles, footnotes,
and exports them to structured CSV files.
"""

import os
import sys
import json
import base64
import csv
import re
from pathlib import Path
from typing import Optional
import anthropic

try:
    import fitz  # pymupdf
except ImportError:
    print("Installing pymupdf...")
    os.system(f"{sys.executable} -m pip install pymupdf -q")
    import fitz


def pdf_page_to_base64(pdf_path: str, page_num: int, dpi: int = 200) -> str:
    """Render a PDF page to a base64-encoded PNG image."""
    doc = fitz.open(pdf_path)
    page = doc[page_num]
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat)
    img_bytes = pix.tobytes("png")
    doc.close()
    return base64.standard_b64encode(img_bytes).decode("utf-8")


def extract_tables_from_page(
    client: anthropic.Anthropic,
    image_b64: str,
    page_num: int,
    pdf_filename: str,
) -> list[dict]:
    """
    Use Claude vision to extract all tables from a single page image.
    Returns a list of table dicts with title, footnotes, headers, and rows.
    """
    prompt = """You are an expert document analyst. Carefully examine this PDF page image and extract ALL tables present, including tables that appear as embedded images or screenshots.

For EACH table found, return a JSON object with these fields:
- "table_index": integer starting from 1 (order of appearance on the page)
- "title": string — the table title or caption (text immediately above the table, e.g. "Table 1: Revenue Summary"). Use null if none found.
- "footnotes": list of strings — any footnotes, notes, or source lines below the table (e.g. "* Excludes depreciation", "Source: Company filings"). Use [] if none.
- "headers": list of strings — the column header names in order
- "rows": list of lists — each inner list is one data row, cells in the same column order as headers. Preserve all cell values exactly as shown. For merged/spanned cells, repeat the value in each spanned position.
- "context_label": string — any section heading or label that contextualizes this table (e.g. "3.2 Financial Results"). Use null if none.

Return your answer as a JSON array containing one object per table. If there are NO tables on this page, return an empty array [].

IMPORTANT:
- Extract EVERY table, even partial tables that continue from a previous page
- For tables in images/screenshots, transcribe the data exactly
- Preserve number formatting (commas, decimals, units, percentages)
- Do not skip any rows or columns
- Return ONLY valid JSON with no additional commentary"""

    collected_text = []

    with client.messages.stream(
        model="claude-opus-4-8",
        max_tokens=8096,
        thinking={"type": "adaptive"},
        messages=[
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
                    {"type": "text", "text": prompt},
                ],
            }
        ],
    ) as stream:
        for text in stream.text_stream:
            collected_text.append(text)

    raw = "".join(collected_text).strip()

    # Extract JSON from response (handle markdown code blocks)
    json_match = re.search(r"```(?:json)?\s*([\s\S]*?)```", raw)
    if json_match:
        raw = json_match.group(1).strip()
    else:
        # Try to find array start
        start = raw.find("[")
        if start != -1:
            raw = raw[start:]

    try:
        tables = json.loads(raw)
    except json.JSONDecodeError:
        print(f"  [warn] Could not parse JSON for page {page_num + 1}, skipping.")
        return []

    if not isinstance(tables, list):
        return []

    # Attach page metadata
    for t in tables:
        t["page_number"] = page_num + 1
        t["source_file"] = pdf_filename

    return tables


def sanitize_filename(name: str, max_len: int = 60) -> str:
    """Create a safe filename from a string."""
    name = re.sub(r'[\\/*?:"<>|]', "_", name)
    name = re.sub(r"\s+", "_", name.strip())
    return name[:max_len] if len(name) > max_len else name


def write_table_csv(table: dict, output_dir: Path, global_index: int) -> str:
    """Write a single table to a CSV file. Returns the file path."""
    title = table.get("title") or f"table_{global_index}"
    safe_title = sanitize_filename(title)
    page = table.get("page_number", 0)
    filename = f"p{page:04d}_{global_index:04d}_{safe_title}.csv"
    filepath = output_dir / filename

    headers = table.get("headers", [])
    rows = table.get("rows", [])

    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)

        # Metadata block
        writer.writerow(["# Source File", table.get("source_file", "")])
        writer.writerow(["# Page Number", page])
        writer.writerow(["# Table Index on Page", table.get("table_index", global_index)])
        writer.writerow(["# Global Table Index", global_index])
        writer.writerow(["# Title", title])
        writer.writerow(["# Context Label", table.get("context_label") or ""])
        footnotes = table.get("footnotes", [])
        for i, fn in enumerate(footnotes, 1):
            writer.writerow([f"# Footnote {i}", fn])
        writer.writerow([])  # blank separator

        # Table data
        if headers:
            writer.writerow(headers)
        for row in rows:
            writer.writerow(row)

    return str(filepath)


def write_combined_csv(all_tables: list[dict], output_dir: Path, pdf_stem: str) -> str:
    """Write all tables into a single combined CSV for easy overview."""
    filepath = output_dir / f"{pdf_stem}_ALL_TABLES.csv"

    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([
            "global_table_index", "source_file", "page_number",
            "table_index_on_page", "title", "context_label",
            "footnotes", "headers", "row_index", "row_data"
        ])

        for g_idx, table in enumerate(all_tables, 1):
            title = table.get("title") or ""
            context = table.get("context_label") or ""
            footnotes = " | ".join(table.get("footnotes", []))
            headers = table.get("headers", [])
            headers_str = " | ".join(str(h) for h in headers)
            rows = table.get("rows", [])

            if not rows:
                # Table with no data rows — write a single summary row
                writer.writerow([
                    g_idx,
                    table.get("source_file", ""),
                    table.get("page_number", ""),
                    table.get("table_index", ""),
                    title, context, footnotes, headers_str, "", ""
                ])
            else:
                for r_idx, row in enumerate(rows, 1):
                    row_str = " | ".join(str(c) for c in row)
                    writer.writerow([
                        g_idx,
                        table.get("source_file", ""),
                        table.get("page_number", ""),
                        table.get("table_index", ""),
                        title, context, footnotes, headers_str, r_idx, row_str
                    ])

    return str(filepath)


def process_pdf(pdf_path: str, output_dir: Path, client: anthropic.Anthropic) -> list[dict]:
    """Process a single PDF file and extract all tables."""
    pdf_path = Path(pdf_path)
    pdf_filename = pdf_path.name
    pdf_stem = pdf_path.stem

    print(f"\n{'='*60}")
    print(f"Processing: {pdf_filename}")

    doc = fitz.open(str(pdf_path))
    total_pages = len(doc)
    doc.close()
    print(f"Total pages: {total_pages}")

    pdf_output_dir = output_dir / sanitize_filename(pdf_stem)
    pdf_output_dir.mkdir(parents=True, exist_ok=True)

    all_tables = []
    global_table_index = 0

    for page_num in range(total_pages):
        print(f"  Page {page_num + 1}/{total_pages}...", end=" ", flush=True)

        image_b64 = pdf_page_to_base64(str(pdf_path), page_num)
        tables = extract_tables_from_page(client, image_b64, page_num, pdf_filename)

        if tables:
            print(f"found {len(tables)} table(s)")
            for table in tables:
                global_table_index += 1
                table["global_table_index"] = global_table_index
                csv_path = write_table_csv(table, pdf_output_dir, global_table_index)
                print(f"    -> Saved: {Path(csv_path).name}")
            all_tables.extend(tables)
        else:
            print("no tables")

    # Write combined CSV
    if all_tables:
        combined_path = write_combined_csv(all_tables, pdf_output_dir, pdf_stem)
        print(f"\n  Combined CSV: {combined_path}")
        print(f"  Total tables extracted: {global_table_index}")
    else:
        print(f"\n  No tables found in {pdf_filename}")

    return all_tables


def main():
    print("PDF Table Extractor — powered by Claude Vision")
    print("=" * 60)

    # Get PDF location
    pdf_input = input("\nEnter path to PDF file or folder containing PDFs: ").strip()
    pdf_input = Path(pdf_input.strip('"').strip("'"))

    if not pdf_input.exists():
        print(f"Error: Path not found: {pdf_input}")
        sys.exit(1)

    if pdf_input.is_file():
        if pdf_input.suffix.lower() != ".pdf":
            print("Error: File must be a PDF.")
            sys.exit(1)
        pdf_files = [pdf_input]
    elif pdf_input.is_dir():
        pdf_files = sorted(pdf_input.glob("*.pdf")) + sorted(pdf_input.glob("*.PDF"))
        if not pdf_files:
            print(f"No PDF files found in: {pdf_input}")
            sys.exit(1)
        print(f"Found {len(pdf_files)} PDF file(s)")
    else:
        print("Error: Invalid path.")
        sys.exit(1)

    # Get output folder
    output_input = input("Enter output folder path for CSV results: ").strip()
    output_dir = Path(output_input.strip('"').strip("'"))
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"Output folder: {output_dir.resolve()}")

    # Initialize Anthropic client
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        api_key = input("Enter your Anthropic API key: ").strip()
    client = anthropic.Anthropic(api_key=api_key)

    # Process each PDF
    total_tables = 0
    for pdf_file in pdf_files:
        tables = process_pdf(str(pdf_file), output_dir, client)
        total_tables += len(tables)

    print(f"\n{'='*60}")
    print(f"Extraction complete.")
    print(f"Total tables extracted across all files: {total_tables}")
    print(f"Results saved to: {output_dir.resolve()}")


if __name__ == "__main__":
    main()
