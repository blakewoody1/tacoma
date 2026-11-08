# Apps Script backend

`Code.gs` is a backup of the Apps Script that runs the Tacoma backend. The live copy is bound to the **Tacoma Maintenance** Google Sheet (Extensions → Apps Script). It also keeps the **Log Tacoma Service** iPhone Shortcut working.

The `KEY` line is a placeholder. The real key lives only in the live script, your Shortcut, and each device's app Settings — never commit it here.

## Install or update the live script
1. Open the Tacoma Maintenance sheet in a browser (on iPad: Safari → aA → Request Desktop Website) → Extensions → Apps Script → Code.gs.
2. Copy your real key from the existing `const KEY = '…';` line.
3. Select all, paste this file's contents, then put your real key back on the `KEY` line.
4. Save.
5. Deploy → Manage deployments → ✎ on the existing deployment → Version: **New version** → Deploy. (Not "New deployment" — that changes the /exec URL your Shortcut uses.)
6. Copy the Web app URL (ends in /exec). In the Tacoma app tap ⚙︎ Settings and paste it, plus your key.

## Keep this backup current
Whenever the live Code.gs changes, update this file too (with the key blanked out) and push.
