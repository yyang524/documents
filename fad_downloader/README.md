# FAD Report Downloader

Mass-downloads reports from the internal OpenText Content Server page
`https://csprod-imfx.opentext.cloud/otcs/cs.exe/app/nodes/1831809`,
walking `Country → Year → Reports → FAD` and saving every report found to
`Knowledge/FAD CD/<Country>/<Year>/`.

## How it stays within the "manual Edge only" constraint

The script does **not** log in, store credentials, or send anonymous
requests. Instead:

1. **You** open Microsoft Edge (via `launch_edge.bat`) and log in to the
   OpenText page yourself, exactly as you do today.
2. The script **attaches to that same Edge window** through Edge's
   built-in debugging port and reuses your live, authenticated session.
3. To navigate folders it calls the very same internal endpoints
   (`/otcs/cs.exe/api/v2/nodes/...`) that the OpenText page itself calls
   in your browser every time you click a folder or a download button —
   it just clicks them for you, with a polite pause (default 0.4 s)
   between requests.

> **Note:** this still automates the clicking. If your organization's
> policy forbids any automated retrieval (not just credential-based
> scraping), please confirm with IT before using it.

## One-time setup

```bat
pip install playwright
```

That's it — no `playwright install` step is needed, because the script
attaches to the Edge you already have.

## Every run

1. Double-click **`launch_edge.bat`**. A separate Edge window opens on the
   OpenText page (it uses its own profile folder, so your normal Edge
   windows are untouched). Log in if prompted. The first time you'll need
   to sign in; afterwards SSO usually signs you in automatically.
2. Leave that window open and run the script from this folder.

### First run — everything since 2008

```bat
python fad_downloader.py --full
```

### Later runs — current year only, duplicates skipped

```bat
python fad_downloader.py
```

Reports already downloaded are remembered in
`Knowledge/FAD CD/.fad_download_manifest.json` and skipped automatically.
Schedule the plain command with Windows Task Scheduler if you want the
update to run on a cadence (Edge must be open and signed in when it runs).

## Useful options

| Option | Meaning |
|---|---|
| `--dry-run` | Show what would be downloaded, download nothing. Try this first. |
| `--year 2023` | Process a single year only. |
| `--full --since 2015` | Full run starting from 2015 instead of 2008. |
| `--out "D:\Reports"` | Save somewhere other than `Knowledge\FAD CD`. |
| `--force` | Re-download even if the manifest already has the report. |
| `--delay 1.0` | Slow down to 1 request/second if you want to be extra gentle. |

## Behavior details

- Output layout: `<out>/<Country>/<Year>/<report file>`, e.g.
  `Knowledge/FAD CD/Saudi Arabia/2024/SAUDI ARABIA Building Fiscal Risk
  Analytical Capability.pdf`.
- Filenames come from the server's download response (same name you'd get
  clicking Download by hand); characters Windows forbids are replaced
  with `_`.
- Folder matching is case-insensitive; `FAD` also matches folders whose
  name merely *starts* with "FAD" (e.g. "FAD CD").
- Countries or years without a `Reports/FAD` path are skipped silently.
- Failed downloads are retried 3 times, then logged; re-running the
  script retries only the failures (successes are in the manifest).
- If your login session expires mid-run, the affected requests fail and
  are reported — just re-run after signing in again.

## Troubleshooting

**"Could not attach to Microsoft Edge"** — Edge wasn't started with the
debugging port. Use `launch_edge.bat` (don't open Edge the normal way for
this) and keep the window open while the script runs.

**Stuck on "Not signed in yet"** — finish the login in the Edge window
the bat file opened (the one showing the OpenText page). The script polls
and continues automatically once you're in.

**Corporate proxy errors from `pip`** — install with your proxy, e.g.
`pip install --proxy http://yourproxy:port playwright`, or ask IT for the
standard pip configuration.
