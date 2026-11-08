# Tacoma

Maintenance app for a 3rd gen Toyota Tacoma, backed by the "Tacoma Maintenance" Google Sheet.

- `index.html` — the whole app (status board, service guides with tools/parts/videos, history + export, service map)
- `sw.js`, `manifest.webmanifest`, `icon-*.png` — lets it install to the Home Screen and open with no signal
- `apps-script/` — backup of the Apps Script that connects the app (and the Log Tacoma Service Shortcut) to the sheet

Hosted with GitHub Pages. Bump `VERSION` in `sw.js` whenever `index.html` changes so phones pick up the new version.
