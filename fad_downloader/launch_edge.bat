@echo off
rem Starts Microsoft Edge with its standard DevTools debugging port enabled,
rem using a dedicated profile folder so your normal Edge windows are not
rem affected. Log in to the OpenText page in the window that opens, keep it
rem open, then run:  python fad_downloader.py

set PROFILE_DIR=%LOCALAPPDATA%\EdgeFADDownloader

start "" msedge.exe ^
  --remote-debugging-port=9222 ^
  --user-data-dir="%PROFILE_DIR%" ^
  "https://csprod-imfx.opentext.cloud/otcs/cs.exe/app/nodes/1831809"

echo Edge is starting. Log in to the OpenText page if prompted,
echo keep the window open, then run:  python fad_downloader.py
pause
