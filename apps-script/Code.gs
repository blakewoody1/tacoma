/**
 * BACKUP COPY. The live copy is in the Tacoma Maintenance sheet → Extensions → Apps Script.
 * The KEY below is a placeholder on purpose — never commit your real key to this public repo.
 *
 * Tacoma — Apps Script for your "Tacoma Maintenance" Google Sheet
 *
 * Two jobs:
 *   1. Your "Log Tacoma Service" iPhone Shortcut keeps working exactly as before
 *      (it sends date, odometer, service, notes, source and key).
 *   2. The Tacoma web app reads your sheet and saves changes through it.
 *
 * After changing anything here:
 *   Save → Deploy → Manage deployments → ✎ (pencil) → Version: New version → Deploy.
 *   (That keeps the same web app URL, so the Shortcut and the app don't need a new address.)
 *
 * Safety rules for the app's requests (so a leaked key can't do much):
 *   - it can only write to the History tab: columns A–E (rows 2 and down, a few rows at a time)
 *     and the app's odometer cells H9:I10
 *   - text starting with "=" is saved as plain text, never as a formula
 *   - new rows are added in one step here, and each one has an ID so a retry never adds it twice
 *   - wrong keys are written to an "Access log" tab; after 30 wrong keys in 10 minutes it stops answering for a while
 */
const KEY = 'CHANGE-ME-to-your-key';

const HISTORY = 'History';

/* ---------- entry points ---------- */
function doGet() {
  return ContentService.createTextOutput('Tacoma backend is running.');
}

function doPost(e) {
  const body = (e && e.postData && e.postData.contents) || '';
  let req = null;
  try { req = JSON.parse(body); } catch (err) { /* not JSON */ }
  if (!req) return ContentService.createTextOutput('Not a Tacoma request.');
  if (req.tool) return json_(api_(req));
  return shortcut_(req);      // the "Log Tacoma Service" Shortcut
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 1. iPhone Shortcut (unchanged behavior) ---------- */
function shortcut_(d) {
  if (d.key !== KEY) { logBad_('Shortcut', d.service); return ContentService.createTextOutput('Wrong key'); }
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(HISTORY);
    const p = String(d.date || '').split('-');
    const day = p.length === 3 ? new Date(p[0], p[1] - 1, p[2]) : new Date(new Date().toDateString());
    const odo = Number(String(d.odometer).replace(/,/g, ''));
    const row = nextRow_(sh);
    sh.getRange(row, 1, 1, 5).setValues([[day, odo, safeText_(String(d.service || '')), safeText_(String(d.notes || '')), safeText_(String(d.source || 'Owner notes'))]]);
    sh.getRange(row, 1).setNumberFormat('M/d/yyyy');
    sh.getRange(row, 2).setNumberFormat('#,##0');
    return ContentService.createTextOutput('Logged ' + d.service + ' at ' + odo.toLocaleString() + ' mi');
  } finally { lock.releaseLock(); }
}

/* ---------- 2. Web app: small read/write API ---------- */
const WRITES_ = ['update_values', 'add_log'];
function api_(req) {
  const cache = CacheService.getScriptCache(), tries = +(cache.get('bad') || 0);
  if (tries >= 30) return { error: 'Too many wrong keys. Try again in 10 minutes.', code: 'needs_reauth' };
  if (req.key !== KEY) { cache.put('bad', String(tries + 1), 600); logBad_('App', req.tool); return { error: 'Wrong key', code: 'needs_reauth' }; }
  const write = WRITES_.indexOf(req.tool) >= 0;
  const lock = write ? LockService.getScriptLock() : null;
  if (lock) lock.waitLock(30000);
  try {
    if (req.tool === 'batch') return { result: { results: ((req.input || {}).calls || []).map(c => run_(c.tool, c.input || {}, true)) } };
    return run_(req.tool, req.input || {}, false);
  } finally {
    if (lock) lock.releaseLock();
  }
}
function run_(tool, i, readOnly) {
  try {
    const ss = SpreadsheetApp.getActive();
    if (i.spreadsheetId && i.spreadsheetId !== ss.getId()) throw new Error('This Apps Script only works with the sheet it is attached to.');
    switch (tool) {
      case 'get_spreadsheet': return { result: getSpreadsheet_(ss) };
      case 'get_values': return { result: getValues_(ss, i) };
      case 'update_values': if (!readOnly) return { result: updateValues_(ss, i) }; break;
      case 'add_log': if (!readOnly) return { result: addLog_(ss, i) }; break;
    }
    return { error: 'Unknown request: ' + tool, code: 'bad_request' };
  } catch (err) {
    return { error: String((err && err.message) || err), code: 'tool_error' };
  }
}

function getSpreadsheet_(ss) {
  return {
    properties: { title: ss.getName(), spreadsheetId: ss.getId(), features: ['batch', 'add_log', 'guard'] },
    sheets: ss.getSheets().map(s => ({ properties: { title: s.getName(), sheetId: s.getSheetId() } }))
  };
}

// Display values (what you see in the sheet), trimmed to the rows that have data
function getValues_(ss, i) {
  const rg = ss.getRange(i.range), sh = rg.getSheet();
  const rows = Math.min(rg.getNumRows(), sh.getLastRow() - rg.getRow() + 1);
  const values = rows > 0 ? sh.getRange(rg.getRow(), rg.getColumn(), rows, rg.getNumColumns()).getDisplayValues().map(r => {
    r = r.slice(); while (r.length && r[r.length - 1] === '') r.pop(); return r;
  }) : [];
  while (values.length && !values[values.length - 1].length) values.pop();
  return { range: i.range, values: values };
}

// Edits and removals from the app (History A–E) and the odometer cells (H9:I10)
function updateValues_(ss, i) {
  const rg = ss.getRange(i.range), sh = rg.getSheet(), vals = i.values || [];
  if (!vals.length) return { status: 'success' };
  const width = Math.max.apply(null, vals.map(r => r.length));
  const c1 = rg.getColumn(), c2 = c1 + width - 1, r1 = rg.getRow(), r2 = r1 + vals.length - 1;
  checkWrite_(sh, c1, c2, r1, r2);
  const tz = ss.getSpreadsheetTimeZone();
  const grid = vals.map(r => {
    const x = r.map(v => toCell_(v, tz));
    while (x.length < width) x.push('');
    return x;
  });
  sh.getRange(r1, c1, grid.length, width).setValues(grid);
  if (c1 === 1) for (let k = 0; k < grid.length; k++) {
    if (grid[k][0] instanceof Date) sh.getRange(r1 + k, 1).setNumberFormat('M/d/yyyy');
    if (typeof grid[k][1] === 'number') sh.getRange(r1 + k, 2).setNumberFormat('#,##0');
  }
  SpreadsheetApp.flush();
  return { updatedRange: i.range, status: 'success' };
}

// Adds one service row at the bottom of History in a single step (an ID makes a retry harmless)
function addLog_(ss, i) {
  if (!i.opId) throw new Error('Missing change ID.');
  if (opDone_(ss, i.opId)) return { status: 'success', duplicate: true };
  const sh = ss.getSheetByName(HISTORY), tz = ss.getSpreadsheetTimeZone();
  const row = (i.row || []).slice(0, 5).map(v => toCell_(v, tz));
  while (row.length < 5) row.push('');
  if (!String(row[2] || '').trim()) throw new Error('Service performed is empty.');
  const r = nextRow_(sh);
  sh.getRange(r, 1, 1, 5).setValues([row]);
  if (row[0] instanceof Date) sh.getRange(r, 1).setNumberFormat('M/d/yyyy');
  if (typeof row[1] === 'number') sh.getRange(r, 2).setNumberFormat('#,##0');
  SpreadsheetApp.flush();
  markOp_(ss, i.opId);
  if (Math.random() < 0.05) pruneOps_();
  return { status: 'success', row: r };
}

/* ---------- helpers ---------- */
// First empty row under the History log (looks only at columns A–E, so the status box in H:I never gets in the way)
function nextRow_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return 2;
  const v = sh.getRange(2, 1, last - 1, 5).getDisplayValues();
  for (let k = v.length - 1; k >= 0; k--) if (v[k].some(x => x !== '')) return k + 3;
  return 2;
}

// The app sends dates as M/D/YYYY (stored as real dates); other text like "3/2026" or "~36,000" stays as typed
function toCell_(v, tz) {
  if (v == null) return '';
  if (typeof v === 'number') return v;
  v = safeText_(String(v));
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(v)) return Utilities.parseDate(v, tz, 'M/d/yyyy');
  return v;
}

// Text that starts with "=" (or +, -, @) is saved as plain text, never run as a formula
function safeText_(v) {
  return /^[=+@]/.test(v) || /^-[^\d]/.test(v) ? "'" + v : v;
}

function checkWrite_(sh, c1, c2, r1, r2) {
  if (sh.getName() !== HISTORY) throw new Error('The app can only change the History tab.');
  const log = c1 >= 1 && c2 <= 5 && r1 >= 2 && (r2 - r1) < 5;
  const odo = c1 === 8 && c2 === 9 && r1 === 9 && r2 === 10;
  if (!log && !odo) throw new Error('That change is outside the cells the app uses.');
}

// Each new row carries an ID; remember the ones already added (kept about 90 days)
function opKey_(ss, id) {
  return 'op:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, ss.getId() + '|' + id));
}
function opDone_(ss, id) { return !!(id && PropertiesService.getScriptProperties().getProperty(opKey_(ss, id))); }
function markOp_(ss, id) { if (id) PropertiesService.getScriptProperties().setProperty(opKey_(ss, id), String(Date.now())); }
function pruneOps_() {
  const props = PropertiesService.getScriptProperties(), all = props.getProperties(), old = Date.now() - 90 * 864e5;
  Object.keys(all).forEach(k => { if (k.indexOf('op:') === 0 && +all[k] < old) props.deleteProperty(k); });
}

// Wrong-key attempts go to an "Access log" tab (created the first time; keeps the last 500 lines)
function logBad_(source, detail) {
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('Access log');
    if (!sh) { sh = ss.insertSheet('Access log'); sh.appendRow(['Time', 'Source', 'What was tried', 'Result']); sh.setFrozenRows(1); }
    if (sh.getLastRow() > 500) sh.deleteRows(2, sh.getLastRow() - 500);
    sh.appendRow([new Date(), source, String(detail || '').slice(0, 60), 'Refused: wrong key']);
  } catch (err) { /* never let logging break a request */ }
}
