/**
 * Print Room Tracker — Google Sheets backend.
 *
 * Paste this whole file into Extensions → Apps Script in the
 * "print-room-tracker" Google Sheet, then Deploy → New deployment →
 * Web app (Execute as: Me, Who has access: Anyone).
 *
 * The script creates two tabs on first use:
 *   Entries — one row per logged task (soft-deleted rows stay, marked deleted)
 *   Config  — the setup (products, steps, targets, people) as JSON in cell A2
 */

const PASSCODE = 'CHANGE-ME';   // change here to rotate; redeploy after changing
const SCRIPT_VERSION = 2;           // 2 = pause / resume
const ENTRY_COLS = ['id','date','start','end','employee','productId','stepId','kind','qty','distributor','printer','note','createdAt','updatedAt','deleted','startLocal','endLocal','minutes','pausedMs','pausedAt'];
const SEED_CONFIG = {"version": 1, "employees": [], "distributors": ["Grove", "Amazon FBA", "Direct / Retail"], "printers": [{"name": "Printer 1", "hoursPerDay": 8}], "productiveHoursPerPerson": 6.5, "workdaysPerWeek": 5, "products": [{"id": "poster", "name": "Posters", "color": 0, "completeStep": "ship", "askDistributor": false, "steps": [{"id": "prep", "name": "Image file prep & queue", "kind": "labor", "est": 30, "share": 1}, {"id": "label", "name": "Print invoice / label", "kind": "labor", "est": 30, "share": 1}, {"id": "print", "name": "Print run", "kind": "machine", "est": null, "share": 1}, {"id": "trim", "name": "Cut & trim", "kind": "labor", "est": 45, "share": 1}, {"id": "roll", "name": "Rolling", "kind": "labor", "est": 30, "share": 1}, {"id": "tube", "name": "Tubing", "kind": "labor", "est": 40, "share": 1}, {"id": "ship", "name": "Pack & ship", "kind": "labor", "est": null, "share": 1}]}, {"id": "vinyl", "name": "Vinyl / Decals", "color": 1, "completeStep": "ship", "askDistributor": false, "steps": [{"id": "prep", "name": "Image file prep & queue", "kind": "labor", "est": 30, "share": 1}, {"id": "label", "name": "Print invoice / label", "kind": "labor", "est": 30, "share": 1}, {"id": "print", "name": "Print run", "kind": "machine", "est": null, "share": 1}, {"id": "trim", "name": "Cut & trim", "kind": "labor", "est": 45, "share": 1}, {"id": "weed", "name": "Weeding (peel excess)", "kind": "labor", "est": 20, "share": 1}, {"id": "roll", "name": "Rolling", "kind": "labor", "est": 30, "share": 1}, {"id": "tube", "name": "Tubing", "kind": "labor", "est": 40, "share": 1}, {"id": "ship", "name": "Pack & ship", "kind": "labor", "est": null, "share": 1}]}, {"id": "canvas", "name": "Canvas", "color": 2, "completeStep": "ship", "askDistributor": false, "steps": [{"id": "prep", "name": "Image file prep & queue", "kind": "labor", "est": 30, "share": 1}, {"id": "label", "name": "Print invoice / label", "kind": "labor", "est": 30, "share": 1}, {"id": "print", "name": "Print run", "kind": "machine", "est": null, "share": 1}, {"id": "trim", "name": "Cut & trim", "kind": "labor", "est": 45, "share": 1}, {"id": "frames", "name": "Cut / build frames", "kind": "labor", "est": 10, "share": 1}, {"id": "stretch", "name": "Stretch canvas", "kind": "labor", "est": 20, "share": 1}, {"id": "foam", "name": "Foam canvases", "kind": "labor", "est": 90, "share": 1}, {"id": "boxes", "name": "Build boxes", "kind": "labor", "est": 6, "share": 1}, {"id": "pack", "name": "Package canvas (box already built)", "kind": "labor", "est": 50, "share": 1}, {"id": "ship", "name": "Ship", "kind": "labor", "est": null, "share": 1}]}, {"id": "wrap_single", "name": "Wrapping Paper \u2013 Singles", "color": 3, "completeStep": "ship", "askDistributor": true, "steps": [{"id": "prep", "name": "Image file prep & queue", "kind": "labor", "est": 30, "share": 1}, {"id": "label", "name": "Print invoice / label", "kind": "labor", "est": 30, "share": 1}, {"id": "print", "name": "Print run", "kind": "machine", "est": null, "share": 1}, {"id": "trim", "name": "Cut & trim", "kind": "labor", "est": 42, "share": 1}, {"id": "rolltube", "name": "Roll & tube", "kind": "labor", "est": 24, "share": 1}, {"id": "sticker", "name": "Sticker", "kind": "labor", "est": null, "share": 1}, {"id": "ship", "name": "Pack & ship", "kind": "labor", "est": null, "share": 1}]}, {"id": "wrap_bulk", "name": "Wrapping Paper \u2013 Bulk", "color": 4, "completeStep": "ship", "askDistributor": true, "defaultDistributor": "Grove", "steps": [{"id": "prep", "name": "Image file prep & queue", "kind": "labor", "est": 30, "share": 1}, {"id": "label", "name": "Print invoice / label", "kind": "labor", "est": 30, "share": 1}, {"id": "print", "name": "Print run", "kind": "machine", "est": null, "share": 1}, {"id": "trim", "name": "Cut & trim", "kind": "labor", "est": 42, "share": 1}, {"id": "rolltube", "name": "Roll & tube", "kind": "labor", "est": 24, "share": 1}, {"id": "sticker", "name": "Sticker", "kind": "labor", "est": null, "share": 1}, {"id": "ship", "name": "Pack & ship", "kind": "labor", "est": null, "share": 1}]}, {"id": "ribbon", "name": "Ribbons", "color": 5, "completeStep": "package", "askDistributor": true, "defaultDistributor": "Grove", "steps": [{"id": "upc", "name": "UPC sticker", "kind": "labor", "est": null, "share": 1}, {"id": "package", "name": "Package", "kind": "labor", "est": null, "share": 1}]}, {"id": "photo", "name": "Photos (ScanCafe prints)", "color": 6, "completeStep": "ship", "askDistributor": false, "steps": [{"id": "pack", "name": "Pack", "kind": "labor", "est": null, "share": 1}, {"id": "ship", "name": "Ship", "kind": "labor", "est": null, "share": 1}]}, {"id": "general", "name": "General / overhead", "color": 7, "overhead": true, "completeStep": null, "askDistributor": false, "steps": [{"id": "maint", "name": "Machine maintenance", "kind": "labor", "perDay": 0.5}, {"id": "clean", "name": "Work space clean up", "kind": "labor", "perDay": 0.3}]}]};

function doGet(e) { return handle_(e.parameter || {}); }
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData && e.postData.contents || '{}'); } catch (err) { return json_({ ok: false, error: 'Bad request body' }); }
  return handle_(body);
}

function handle_(req) {
  try {
    if (req.key !== PASSCODE) return json_({ ok: false, error: 'bad_passcode' });
    switch (req.action) {
      case 'load':       return json_(load_(req.since));
      case 'changes':    return json_(changes_(Number(req.after) || 0));
      case 'upsert':     return json_(withLock_(() => upsert_(req.entry)));
      case 'stop':       return json_(withLock_(() => stop_(req.id, Number(req.end), req.qty)));
      case 'delete':     return json_(withLock_(() => patch_(req.id, { deleted: true })));
      case 'pause':      return json_(withLock_(() => pause_(req.id, Number(req.at))));
      case 'resume':     return json_(withLock_(() => resume_(req.id, Number(req.at))));
      case 'saveConfig': return json_(withLock_(() => saveConfig_(req.config)));
      case 'ping':       return json_({ ok: true, serverTime: Date.now(), scriptVersion: SCRIPT_VERSION });
      default:           return json_({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function withLock_(fn) { const lock = LockService.getScriptLock(); lock.waitLock(20000); try { return fn(); } finally { lock.releaseLock(); } }

function entriesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Entries');
  if (!sh) {
    sh = ss.insertSheet('Entries');
    sh.getRange(1, 1, 1, ENTRY_COLS.length).setValues([ENTRY_COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('B:B').setNumberFormat('@');
  } else if (sh.getLastColumn() < ENTRY_COLS.length) {
    sh.getRange(1, 1, 1, ENTRY_COLS.length).setValues([ENTRY_COLS]).setFontWeight('bold');  // add new columns to an older sheet
  }
  return sh;
}
function configSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Config');
  if (!sh) {
    sh = ss.insertSheet('Config');
    sh.getRange('A1').setValue('Setup JSON (edit from the tracker\'s Setup tab, not here)').setFontWeight('bold');
    sh.getRange('A2').setValue(JSON.stringify(SEED_CONFIG));
    sh.getRange('B1').setValue('updatedAt');
    sh.getRange('B2').setValue(Date.now());
  }
  return sh;
}
function getConfig_() {
  const sh = configSheet_();
  return { config: JSON.parse(sh.getRange('A2').getValue() || 'null'), configUpdatedAt: Number(sh.getRange('B2').getValue()) || 0 };
}
function saveConfig_(config) {
  if (!config || !Array.isArray(config.products)) throw new Error('Setup is missing products');
  const sh = configSheet_();
  const now = Date.now();
  sh.getRange('A2').setValue(JSON.stringify(config));
  sh.getRange('B2').setValue(now);
  return { ok: true, configUpdatedAt: now };
}

function rowToEntry_(r) {
  const e = {};
  ENTRY_COLS.forEach((c, i) => { e[c] = r[i]; });
  e.id = String(e.id);
  e.date = typeof e.date === 'object' && e.date ? Utilities.formatDate(e.date, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(e.date);
  ['start','end','createdAt','updatedAt','pausedAt'].forEach(k => { e[k] = e[k] === '' || e[k] == null ? null : Number(e[k]); });
  e.pausedMs = Number(e.pausedMs) || 0;
  e.qty = e.qty === '' ? null : Number(e.qty);
  e.deleted = e.deleted === true || e.deleted === 'TRUE';
  ['employee','distributor','printer','note'].forEach(k => { if (e[k] === '') e[k] = null; });
  delete e.startLocal; delete e.endLocal; delete e.minutes;
  return e;
}
function allRows_() {
  const sh = entriesSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, ENTRY_COLS.length).getValues();
}
function load_(since) {
  const rows = allRows_().map(rowToEntry_).filter(e => !e.deleted && (!since || e.date >= since));
  const c = getConfig_();
  return { ok: true, entries: rows, config: c.config, configUpdatedAt: c.configUpdatedAt, serverTime: Date.now(), scriptVersion: SCRIPT_VERSION };
}
function changes_(after) {
  const cutoff = after - 10000; // overlap so nothing slips between polls
  const rows = allRows_().map(rowToEntry_).filter(e => e.updatedAt && e.updatedAt > cutoff);
  const c = getConfig_();
  return { ok: true, entries: rows, config: c.configUpdatedAt > cutoff ? c.config : null, configUpdatedAt: c.configUpdatedAt, serverTime: Date.now() };
}

function localTime_(ms) { return ms ? Utilities.formatDate(new Date(Number(ms)), Session.getScriptTimeZone(), 'h:mm a') : ''; }
function rowFor_(e) {
  const mins = e.start && e.end ? Math.round((Number(e.end) - Number(e.start) - (Number(e.pausedMs) || 0)) / 600) / 100 : '';
  const v = Object.assign({}, e, { startLocal: localTime_(e.start), endLocal: localTime_(e.end), minutes: mins });
  return ENTRY_COLS.map(c => (v[c] === null || v[c] === undefined) ? '' : v[c]);
}
function findRow_(id) {
  const sh = entriesSheet_();
  if (sh.getLastRow() < 2) return 0;
  const hit = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}
function upsert_(entry) {
  if (!entry || !entry.productId || !entry.stepId || !entry.start || !entry.date) throw new Error('Entry is missing product, step, date or start time');
  const sh = entriesSheet_();
  const now = Date.now();
  const e = Object.assign({}, entry);
  e.id = e.id || Utilities.getUuid().replace(/-/g, '').slice(0, 20);
  e.createdAt = e.createdAt || now;
  e.updatedAt = now;
  e.deleted = false;
  e.note = e.note ? String(e.note).slice(0, 200) : '';
  const row = findRow_(e.id);
  if (row) sh.getRange(row, 1, 1, ENTRY_COLS.length).setValues([rowFor_(e)]);
  else sh.appendRow(rowFor_(e));
  return { ok: true, entry: e };
}
function patch_(id, fields) {
  const row = findRow_(id);
  if (!row) throw new Error('That entry no longer exists');
  const sh = entriesSheet_();
  const e = rowToEntry_(sh.getRange(row, 1, 1, ENTRY_COLS.length).getValues()[0]);
  Object.assign(e, fields, { updatedAt: Date.now() });
  sh.getRange(row, 1, 1, ENTRY_COLS.length).setValues([rowFor_(e)]);
  return { ok: true, entry: e };
}

function stop_(id, end, qty) {
  const row = findRow_(id); if (!row) throw new Error('That entry no longer exists');
  const e = rowToEntry_(entriesSheet_().getRange(row, 1, 1, ENTRY_COLS.length).getValues()[0]);
  const fields = { end: end, qty: qty === '' || qty == null ? '' : Number(qty) };
  if (e.pausedAt) { fields.pausedMs = e.pausedMs + Math.max(0, end - e.pausedAt); fields.pausedAt = null; }  // stopped while paused
  return patch_(id, fields);
}
function pause_(id, at) {
  const row = findRow_(id); if (!row) throw new Error('That entry no longer exists');
  const e = rowToEntry_(entriesSheet_().getRange(row, 1, 1, ENTRY_COLS.length).getValues()[0]);
  if (e.end) throw new Error('That timer is already stopped');
  if (e.pausedAt) return { ok: true, entry: e };       // already paused
  return patch_(id, { pausedAt: at || Date.now() });
}
function resume_(id, at) {
  const row = findRow_(id); if (!row) throw new Error('That entry no longer exists');
  const e = rowToEntry_(entriesSheet_().getRange(row, 1, 1, ENTRY_COLS.length).getValues()[0]);
  if (!e.pausedAt) return { ok: true, entry: e };      // not paused
  return patch_(id, { pausedMs: e.pausedMs + Math.max(0, (at || Date.now()) - e.pausedAt), pausedAt: null });
}

/** Run once from the editor (select setup, press Run) to create the tabs and grant access. */
function setup() { entriesSheet_(); configSheet_(); }
