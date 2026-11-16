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
const SCRIPT_VERSION = 3;           // 2 = pause / resume, 3 = emailed forecast reports
const MAX_RECIPIENTS = 25;          // per report
const MAX_REPORTS_PER_DAY = 20;     // guards the Gmail sending quota
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
      case 'emailReport': return json_(withLock_(() => emailReport_(req.report)));
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

// ---------- emailed forecast report ----------
const EMAIL_RE = /^[^\s@<>(),;:"]+@[^\s@<>(),;:"]+\.[^\s@<>(),;:"]+$/;
function esc_(v) { return String(v == null ? '' : v).replace(/[&<>"']/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c])); }
function num_(v, dp) { const n = Number(v); if (!isFinite(n)) return '—'; return n.toLocaleString('en-US', { minimumFractionDigits: dp || 0, maximumFractionDigits: dp || 0 }); }
function hrs_(v) { return num_(Number(v) || 0, 1); }
function str_(v, max) { return String(v == null ? '' : v).slice(0, max || 120); }

function emailReport_(r) {
  if (!r || typeof r !== 'object') throw new Error('The report was empty.');
  // recipients come only from the saved setup, never from the request
  const cfg = getConfig_().config || {};
  const to = (cfg.reportEmails || []).map(e => String(e).trim().toLowerCase()).filter(e => EMAIL_RE.test(e));
  const list = [...new Set(to)].slice(0, MAX_RECIPIENTS);
  if (!list.length) throw new Error('No one is on the mailing list yet. Add addresses on the Setup tab.');

  const props = PropertiesService.getScriptProperties();
  const dayKey = 'reports-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const sentToday = Number(props.getProperty(dayKey)) || 0;
  if (sentToday >= MAX_REPORTS_PER_DAY) throw new Error('The daily limit of ' + MAX_REPORTS_PER_DAY + ' reports has been reached. Try again tomorrow.');
  if (MailApp.getRemainingDailyQuota() < list.length) throw new Error('Gmail\'s daily sending limit is used up. Try again tomorrow.');

  const period = ['day','week','month'].indexOf(r.period) >= 0 ? r.period : 'day';
  const label = str_(r.label, 60);
  const products = (Array.isArray(r.products) ? r.products : []).slice(0, 30).map(p => ({ name: str_(p.name, 60), units: Number(p.units) || 0 })).filter(p => p.units > 0);
  const steps = (Array.isArray(r.steps) ? r.steps : []).slice(0, 120).map(x => ({ product: str_(x.product, 60), step: str_(x.step, 80), units: x.units == null ? null : Number(x.units), hours: Number(x.hours) || 0, printer: !!x.printer }));
  const labor = Number(r.laborHours) || 0, people = Number(r.peopleNeeded) || 0, printer = Number(r.printerHours) || 0, cap = Number(r.printerCapacity) || 0;
  const hpp = Number(r.hoursPerPerson) || 0, missing = Number(r.missingSteps) || 0;
  const sentBy = str_(r.sentBy, 60), note = str_(r.note, 500), pageUrl = /^https:\/\/[a-z0-9.-]+\.github\.io\//i.test(String(r.pageUrl || '')) ? String(r.pageUrl).slice(0, 200) : '';

  const subject = 'Print room forecast, ' + label + ': ' + hrs_(labor) + ' labor hours';
  const ink = '#132029', muted = '#5a6873', line = '#d6dee5', accent = '#0b6f93', bad = '#b0281f';
  const tile = (v, l, s) => '<td style="padding:10px 14px;border-left:3px solid ' + accent + ';vertical-align:top"><div style="font:600 11px Arial,sans-serif;letter-spacing:.06em;text-transform:uppercase;color:' + muted + '">' + l + '</div><div style="font:700 26px Arial,sans-serif;color:' + ink + '">' + v + '</div><div style="font:12px Arial,sans-serif;color:' + muted + '">' + s + '</div></td>';
  const th = t => '<th style="text-align:left;padding:6px 8px;border-bottom:1px solid ' + line + ';font:600 11px Arial,sans-serif;letter-spacing:.05em;text-transform:uppercase;color:' + muted + '">' + t + '</th>';
  const thn = t => th(t).replace('text-align:left', 'text-align:right');
  const td = t => '<td style="padding:6px 8px;border-bottom:1px solid ' + line + ';font:14px Arial,sans-serif;color:' + ink + '">' + t + '</td>';
  const tdn = t => td(t).replace('font:14px', 'text-align:right;font:13px Consolas,Menlo,monospace;font-size:13px');

  let html = '<div style="max-width:640px;font-family:Arial,sans-serif;color:' + ink + '">';
  html += '<div style="background:#132a38;color:#e9f1f6;padding:14px 18px;font:700 18px Arial,sans-serif;letter-spacing:.03em;text-transform:uppercase">Print room forecast</div>';
  html += '<p style="font:15px Arial,sans-serif;margin:16px 0 4px"><b>' + esc_(label) + '</b>' + (sentBy ? ' · sent by ' + esc_(sentBy) : '') + '</p>';
  if (note) html += '<p style="font:14px Arial,sans-serif;margin:8px 0;padding:10px 12px;background:#eef2f5;border-radius:4px">' + esc_(note).replace(/\n/g, '<br>') + '</p>';
  html += '<table role="presentation" cellspacing="0" cellpadding="0" style="margin:14px 0;border-collapse:collapse"><tr>' +
    tile(hrs_(labor), 'Labor hours', 'incl. overhead') +
    tile(num_(people, 1), 'People needed', hpp ? 'at ' + num_(hpp, 2).replace(/\.?0+$/, '') + ' productive hrs each' : '') +
    tile(printer ? hrs_(printer) : '—', 'Printer hours', cap ? 'of ' + hrs_(cap) + ' available' : '') + '</tr></table>';
  if (cap && printer > cap) html += '<p style="font:14px Arial,sans-serif;color:' + bad + '">Printers are over capacity for this volume. The print queue, not staffing, sets the pace.</p>';
  if (missing) html += '<p style="font:13px Arial,sans-serif;color:' + muted + '">' + missing + ' step' + (missing > 1 ? 's have' : ' has') + ' no target yet and ' + (missing > 1 ? 'are' : 'is') + ' left out of these hours.</p>';
  if (products.length) {
    html += '<h3 style="font:600 13px Arial,sans-serif;letter-spacing:.06em;text-transform:uppercase;color:' + muted + ';margin:18px 0 6px">Planned units</h3><table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%"><tr>' + th('Product') + thn('Units') + '</tr>';
    products.forEach(p => { html += '<tr>' + td(esc_(p.name)) + tdn(num_(p.units)) + '</tr>'; });
    html += '</table>';
  }
  if (steps.length) {
    html += '<h3 style="font:600 13px Arial,sans-serif;letter-spacing:.06em;text-transform:uppercase;color:' + muted + ';margin:18px 0 6px">Hours by step</h3><table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%"><tr>' + th('Product') + th('Step') + thn('Units') + thn('Hours') + '</tr>';
    steps.forEach(x => { html += '<tr>' + td(esc_(x.product)) + td(esc_(x.step) + (x.printer ? ' <span style="color:' + accent + ';font-size:12px">(printer)</span>' : '')) + tdn(x.units == null ? '—' : num_(x.units)) + tdn(hrs_(x.hours)) + '</tr>'; });
    html += '</table>';
  }
  html += '<p style="font:12px Arial,sans-serif;color:' + muted + ';margin-top:18px">Printer hours run alongside hands-on work and are not part of labor hours.' + (pageUrl ? ' <a href="' + esc_(pageUrl) + '" style="color:' + accent + '">Open the tracker</a>' : '') + '</p></div>';

  let text = 'Print room forecast, ' + label + (sentBy ? ' (sent by ' + sentBy + ')' : '') + '\n\n';
  if (note) text += note + '\n\n';
  text += 'Labor hours: ' + hrs_(labor) + ' (incl. overhead)\nPeople needed: ' + num_(people, 1) + '\nPrinter hours: ' + (printer ? hrs_(printer) : '—') + (cap ? ' of ' + hrs_(cap) + ' available' : '') + '\n\n';
  if (products.length) text += 'Planned units\n' + products.map(p => '  ' + p.name + ': ' + num_(p.units)).join('\n') + '\n\n';
  if (steps.length) text += 'Hours by step\n' + steps.map(x => '  ' + x.product + ' · ' + x.step + ': ' + hrs_(x.hours) + ' h').join('\n') + '\n';

  MailApp.sendEmail({ to: list.join(','), subject: subject, htmlBody: html, body: text, name: 'Print Room Tracker' });
  props.setProperty(dayKey, String(sentToday + 1));

  // keep a record of what went out
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let log = ss.getSheetByName('Reports');
  if (!log) { log = ss.insertSheet('Reports'); log.getRange(1, 1, 1, 8).setValues([['sentAt','report','period','laborHours','peopleNeeded','printerHours','sentBy','recipients']]).setFontWeight('bold'); log.setFrozenRows(1); }
  log.appendRow([new Date(), label, period, Math.round(labor * 10) / 10, Math.round(people * 10) / 10, Math.round(printer * 10) / 10, sentBy, list.join(', ')]);
  return { ok: true, sentTo: list.length };
}

/** Run once from the editor after pasting this version, to allow the script to send email as you. */
function authorizeEmail() { Logger.log('Emails left today: ' + MailApp.getRemainingDailyQuota()); }

/** Run once from the editor (select setup, press Run) to create the tabs and grant access. */
function setup() { entriesSheet_(); configSheet_(); }
