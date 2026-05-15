# Tutor Tracker

Simple single-page PWA for logging tutoring sessions, tracking completion/payment status, and syncing entries to a Google Apps Script webhook.

## Features
- Session logging with per-student hourly/session rate.
- Payment tracking (`paid` / `owed`) and completion status.
- Local persistence in browser `localStorage`.
- Bulk import support for quick backfilling.
- Optional sync to Google Sheets via Apps Script webhook.
- Offline support through a service worker cache.
- PWA installability via `manifest.json`.
- URL-driven tab state (`?tab=log|history|students`) so home-screen shortcuts open the expected view.
- Browser back/forward navigation keeps tab and URL in sync.

## Files
- `index.html` — app UI and logic (React + Babel in-browser).
- `service-worker.js` — cache and offline/fetch behavior.
- `manifest.json` — PWA metadata and app shortcuts.

## Run locally
Because this app registers a service worker, serve it over HTTP locally:

```bash
python -m http.server 8000
```

Then open `http://localhost:8000`.

## Notes
- Update `window.SHEETS_WEBHOOK_URL` in `index.html` if your Apps Script URL changes.
- If you change cached assets, bump `CACHE_NAME` in `service-worker.js`.
