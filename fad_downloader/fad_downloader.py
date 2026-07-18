#!/usr/bin/env python3
"""
FAD report mass-downloader for OpenText Content Server.

Walks the folder tree under the root node:

    <root>/<Country>/<Year>/Reports/FAD/<report documents>

and saves every document found in a FAD folder to:

    <output dir>/<Country>/<Year>/<report filename>

Authentication strategy
-----------------------
This script never logs in on its own and never talks to the server
anonymously. You open Microsoft Edge yourself (via launch_edge.bat, which
simply starts Edge with its standard DevTools debugging port enabled), log
in to the OpenText page exactly as you always do, and leave the window
open. The script then *attaches* to that already-authenticated Edge
session and issues the same REST calls (/otcs/cs.exe/api/v2/nodes/...)
that the Smart UI page itself makes every time you click a folder — using
your own cookies, from your own browser. No credentials are ever seen or
stored by the script.

Usage
-----
First run (everything since 2008):

    python fad_downloader.py --full

Later runs (current year only, previously downloaded reports skipped):

    python fad_downloader.py

Other useful flags:

    --since 2015          full run but starting at 2015 instead of 2008
    --year 2023           only process one specific year
    --out "D:/somewhere"  change the output folder
    --dry-run             list what would be downloaded, download nothing
    --force               re-download even if the manifest says we have it
    --delay 0.5           seconds to pause between server requests

Requirements:  pip install playwright
(no "playwright install" needed - we attach to your installed Edge,
nothing is downloaded)
"""

import argparse
import json
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    sys.exit(
        "Playwright is not installed. Run:\n"
        "    pip install playwright\n"
        "(no 'playwright install' step is needed for this script)"
    )

# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------
BASE_URL = "https://csprod-imfx.opentext.cloud/otcs/cs.exe"
ROOT_NODE_ID = 1831809
CDP_URL = "http://localhost:9222"          # Edge remote-debugging endpoint
DEFAULT_OUTPUT_DIR = Path("Knowledge") / "FAD CD"
MANIFEST_NAME = ".fad_download_manifest.json"

FIRST_YEAR_DEFAULT = 2008
PAGE_SIZE = 200                            # OTCS API page size for listings

# OpenText Content Server node subtypes
SUBTYPE_FOLDER = 0
SUBTYPE_DOCUMENT = 144

YEAR_RE = re.compile(r"^(19|20)\d{2}$")
ILLEGAL_FS_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def log(msg: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


def sanitize(name: str) -> str:
    """Make a node name safe to use as a Windows file/folder name."""
    name = ILLEGAL_FS_CHARS.sub("_", name).strip().rstrip(". ")
    return name or "unnamed"


# ----------------------------------------------------------------------------
# Manifest (remembers what was already downloaded, for incremental runs)
# ----------------------------------------------------------------------------
class Manifest:
    def __init__(self, path: Path):
        self.path = path
        self.data = {"downloaded": {}}
        if path.exists():
            try:
                self.data = json.loads(path.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError):
                log(f"WARNING: could not read manifest {path}; starting a new one")
        self.data.setdefault("downloaded", {})

    def has(self, node_id: int) -> bool:
        return str(node_id) in self.data["downloaded"]

    def add(self, node_id: int, info: dict) -> None:
        self.data["downloaded"][str(node_id)] = info
        self.save()

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(
            json.dumps(self.data, indent=2, ensure_ascii=False), encoding="utf-8"
        )


# ----------------------------------------------------------------------------
# OpenText Content Server access through the attached Edge session
# ----------------------------------------------------------------------------
class ContentServer:
    def __init__(self, request_context, delay: float):
        self.request = request_context   # playwright APIRequestContext (shares Edge cookies)
        self.delay = delay

    def _get(self, url: str, max_attempts: int = 3):
        last_error = None
        for attempt in range(1, max_attempts + 1):
            try:
                response = self.request.get(url, timeout=120_000)
                if response.ok:
                    return response
                last_error = f"HTTP {response.status}"
            except Exception as exc:  # network hiccups, timeouts
                last_error = str(exc)
            if attempt < max_attempts:
                time.sleep(2 * attempt)
        raise RuntimeError(f"GET {url} failed after {max_attempts} attempts: {last_error}")

    def is_authenticated(self) -> bool:
        try:
            response = self.request.get(
                f"{BASE_URL}/api/v2/nodes/{ROOT_NODE_ID}", timeout=30_000
            )
        except Exception:
            return False
        if not response.ok:
            return False
        # An expired session often redirects to an HTML login page with status 200,
        # so insist on a JSON body as proof we are really signed in.
        try:
            response.json()
            return True
        except Exception:
            return False

    def children(self, node_id: int) -> list:
        """List all children of a node (follows API paging)."""
        results, page = [], 1
        while True:
            url = (
                f"{BASE_URL}/api/v2/nodes/{node_id}/nodes"
                f"?limit={PAGE_SIZE}&page={page}"
                "&fields=properties{id,name,type,size,modify_date}"
            )
            payload = self._get(url).json()
            for item in payload.get("results", []):
                properties = item.get("data", {}).get("properties", {})
                if properties:
                    results.append(properties)
            paging = payload.get("collection", {}).get("paging", {})
            if page >= paging.get("page_total", 1):
                break
            page += 1
            time.sleep(self.delay)
        time.sleep(self.delay)
        return results

    def download(self, node_id: int, fallback_name: str, target_dir: Path) -> Path:
        response = self._get(f"{BASE_URL}/api/v2/nodes/{node_id}/content")
        filename = _filename_from_headers(response.headers) or sanitize(fallback_name)
        target_dir.mkdir(parents=True, exist_ok=True)
        target = target_dir / sanitize(filename)
        target.write_bytes(response.body())
        time.sleep(self.delay)
        return target


def _filename_from_headers(headers: dict) -> str | None:
    disposition = headers.get("content-disposition", "")
    match = re.search(r"filename\*=(?:UTF-8''|utf-8'')([^;]+)", disposition)
    if match:
        from urllib.parse import unquote
        return unquote(match.group(1).strip().strip('"'))
    match = re.search(r'filename="?([^";]+)"?', disposition)
    if match:
        return match.group(1).strip()
    return None


def find_subfolder(entries: list, wanted: str) -> dict | None:
    """Find a child folder by name: exact case-insensitive match first,
    then a folder whose name starts with the wanted name (e.g. 'FAD CD')."""
    wanted_lower = wanted.lower()
    folders = [e for e in entries if e.get("type") == SUBTYPE_FOLDER]
    for entry in folders:
        if entry.get("name", "").strip().lower() == wanted_lower:
            return entry
    for entry in folders:
        if entry.get("name", "").strip().lower().startswith(wanted_lower):
            return entry
    return None


# ----------------------------------------------------------------------------
# Main traversal
# ----------------------------------------------------------------------------
def run(cs: ContentServer, years_wanted, output_dir: Path, manifest: Manifest,
        dry_run: bool, force: bool) -> None:
    stats = {"downloaded": 0, "skipped": 0, "errors": 0}

    countries = [c for c in cs.children(ROOT_NODE_ID) if c.get("type") == SUBTYPE_FOLDER]
    log(f"Found {len(countries)} country folders under root node {ROOT_NODE_ID}")

    for country in sorted(countries, key=lambda c: c.get("name", "")):
        country_name = country["name"].strip()
        try:
            year_entries = cs.children(country["id"])
        except RuntimeError as exc:
            log(f"ERROR listing {country_name}: {exc}")
            stats["errors"] += 1
            continue

        for year_entry in year_entries:
            if year_entry.get("type") != SUBTYPE_FOLDER:
                continue
            year_name = year_entry.get("name", "").strip()
            if not YEAR_RE.match(year_name) or int(year_name) not in years_wanted:
                continue

            try:
                year_children = cs.children(year_entry["id"])
                reports = find_subfolder(year_children, "Reports")
                if not reports:
                    continue
                fad = find_subfolder(cs.children(reports["id"]), "FAD")
                if not fad:
                    continue
                documents = [d for d in cs.children(fad["id"])
                             if d.get("type") == SUBTYPE_DOCUMENT]
            except RuntimeError as exc:
                log(f"ERROR walking {country_name}/{year_name}: {exc}")
                stats["errors"] += 1
                continue

            if not documents:
                continue
            log(f"{country_name}/{year_name}/Reports/FAD: {len(documents)} document(s)")

            target_dir = output_dir / sanitize(country_name) / sanitize(year_name)
            for document in documents:
                doc_id, doc_name = document["id"], document.get("name", str(document["id"]))
                if manifest.has(doc_id) and not force:
                    log(f"  skip (already downloaded): {doc_name}")
                    stats["skipped"] += 1
                    continue
                if dry_run:
                    log(f"  would download: {doc_name} -> {target_dir}")
                    stats["downloaded"] += 1
                    continue
                try:
                    saved = cs.download(doc_id, doc_name, target_dir)
                except RuntimeError as exc:
                    log(f"  ERROR downloading {doc_name}: {exc}")
                    stats["errors"] += 1
                    continue
                manifest.add(doc_id, {
                    "name": doc_name,
                    "country": country_name,
                    "year": year_name,
                    "path": str(saved),
                    "downloaded_at": datetime.now(timezone.utc).isoformat(),
                    "modify_date": document.get("modify_date"),
                })
                log(f"  downloaded: {saved}")
                stats["downloaded"] += 1

    verb = "would download" if dry_run else "downloaded"
    log(f"Done. {stats['downloaded']} {verb}, "
        f"{stats['skipped']} skipped, {stats['errors']} errors.")
    if stats["errors"]:
        log("Some items failed - re-running the script will retry them "
            "(successful downloads are remembered and skipped).")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Download FAD reports from OpenText Content Server "
                    "through your own logged-in Edge session.")
    parser.add_argument("--full", action="store_true",
                        help=f"download all years since {FIRST_YEAR_DEFAULT} "
                             "(first run); default is current year only")
    parser.add_argument("--since", type=int, default=FIRST_YEAR_DEFAULT,
                        help="with --full: earliest year to include")
    parser.add_argument("--year", type=int,
                        help="process only this single year")
    parser.add_argument("--out", type=Path, default=DEFAULT_OUTPUT_DIR,
                        help=f'output folder (default: "{DEFAULT_OUTPUT_DIR}")')
    parser.add_argument("--cdp-url", default=CDP_URL,
                        help="Edge remote-debugging URL (default: %(default)s)")
    parser.add_argument("--delay", type=float, default=0.4,
                        help="pause between requests in seconds (default: 0.4)")
    parser.add_argument("--dry-run", action="store_true",
                        help="show what would be downloaded without downloading")
    parser.add_argument("--force", action="store_true",
                        help="re-download reports even if already in the manifest")
    args = parser.parse_args()

    current_year = datetime.now().year
    if args.year:
        years_wanted = {args.year}
    elif args.full:
        years_wanted = set(range(args.since, current_year + 1))
    else:
        years_wanted = {current_year}
    log(f"Years to process: {min(years_wanted)}-{max(years_wanted)}"
        if len(years_wanted) > 1 else f"Year to process: {min(years_wanted)}")

    manifest = Manifest(args.out / MANIFEST_NAME)

    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp(args.cdp_url)
        except Exception:
            sys.exit(
                "\nCould not attach to Microsoft Edge.\n\n"
                "Please start Edge with its debugging port enabled first:\n"
                "  1. Double-click launch_edge.bat (in this folder), and\n"
                "  2. log in to the OpenText page in the window that opens,\n"
                "  3. leave that window open and run this script again.\n"
            )
        context = browser.contexts[0] if browser.contexts else browser.new_context()
        cs = ContentServer(context.request, delay=args.delay)

        if not cs.is_authenticated():
            log("Not signed in yet. Please finish logging in to the OpenText "
                "page in the Edge window that is open...")
            log("(waiting up to 5 minutes - the script continues automatically)")
            deadline = time.time() + 300
            while time.time() < deadline:
                time.sleep(5)
                if cs.is_authenticated():
                    break
            else:
                sys.exit("Still not signed in after 5 minutes; giving up. "
                         "Log in in the Edge window, then re-run the script.")
        log("Attached to your Edge session - signed in OK.")

        run(cs, years_wanted, args.out, manifest,
            dry_run=args.dry_run, force=args.force)


if __name__ == "__main__":
    main()
