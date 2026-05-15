# Tutor Tracker PWA

This project is a fully installable Progressive Web App (PWA) that logs tutoring sessions locally and syncs to a Google Apps Script backend.

## What this setup does
- **Sheet1** = live log (always updated by app activity).
- **Sheet2** = archive (append-only for completed sessions only).
- **Sheet2 backup column** = JSON string of the whole session for safe copy/paste re-import.

## Files
- `Code.gs` — Google Apps Script webhook backend.
- `index.html` — Frontend UI + sync logic.
- `manifest.json` — PWA metadata.
- `service-worker.js` — Offline/app-shell caching.

## Configure
1. Deploy your Apps Script as a **Web app**.
2. Copy the web app URL.
3. In `index.html`, set:
   - `window.SHEETS_WEBHOOK_URL = "<your web app url>"`

## Local run
```bash
python -m http.server 8000
```
Open `http://localhost:8000`.
