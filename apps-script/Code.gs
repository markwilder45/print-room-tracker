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
const SCRIPT_VERSION = 5;           // 2 = pause / resume, 3 = emailed forecast reports, 4 = oldest order dates, 5 = yesterday's results + saved plans
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
      case 'getPlan':     return json_(getPlan_(String(req.date || '')));
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
  const MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const dateLabel_ = v => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(v || '')); return m ? MON[Number(m[2]) - 1] + ' ' + Number(m[3]) : ''; };
  const products = (Array.isArray(r.products) ? r.products : []).slice(0, 30).map(p => {
    const oldest = dateLabel_(p.oldest), days = oldest && isFinite(Number(p.daysWaiting)) ? Math.max(0, Math.round(Number(p.daysWaiting))) : null;
    return { name: str_(p.name, 60), units: Number(p.units) || 0, oldest: oldest, days: days };
  }).filter(p => p.units > 0 || p.oldest);
  const waiting = products.filter(p => p.days != null).sort((a, b) => b.days - a.days);
  const behind = waiting[0] || null;   // the product whose oldest order has waited longest
  const steps = (Array.isArray(r.steps) ? r.steps : []).slice(0, 120).map(x => ({ product: str_(x.product, 60), step: str_(x.step, 80), units: x.units == null ? null : Number(x.units), hours: Number(x.hours) || 0, printer: !!x.printer }));
  const labor = Number(r.laborHours) || 0, people = Number(r.peopleNeeded) || 0, printer = Number(r.printerHours) || 0, cap = Number(r.printerCapacity) || 0;
  const hpp = Number(r.hoursPerPerson) || 0, missing = Number(r.missingSteps) || 0;
  const res = parseResults_(r.results);
  const sentBy = str_(r.sentBy, 60), note = str_(r.note, 500), pageUrl = /^https:\/\/[a-z0-9.-]+\.github\.io\//i.test(String(r.pageUrl || '')) ? String(r.pageUrl).slice(0, 200) : '';

  const pct_ = v => v == null ? '—' : Math.round(v * 100) + '%';
  const subject = (res && res.productivity != null
      ? 'Print room, ' + label + ': yesterday ' + pct_(res.productivity) + ' productivity (goal ' + res.goal + '%) · today ' + hrs_(labor) + ' labor hours'
      : 'Print room forecast, ' + label + ': ' + hrs_(labor) + ' labor hours')
    + (behind ? ' · oldest order ' + behind.days + (behind.days === 1 ? ' day' : ' days') : '');
  const daysTxt = d => d == null ? '' : d === 0 ? 'today' : d + (d === 1 ? ' day' : ' days');
  const ink = '#132029', muted = '#5a6873', line = '#d6dee5', accent = '#0b6f93', bad = '#b0281f';
  const tile = (v, l, s) => '<td style="padding:10px 14px;border-left:3px solid ' + accent + ';vertical-align:top"><div style="font:600 11px Arial,sans-serif;letter-spacing:.06em;text-transform:uppercase;color:' + muted + '">' + l + '</div><div style="font:700 26px Arial,sans-serif;color:' + ink + '">' + v + '</div><div style="font:12px Arial,sans-serif;color:' + muted + '">' + s + '</div></td>';
  const th = t => '<th style="text-align:left;padding:6px 8px;border-bottom:1px solid ' + line + ';font:600 11px Arial,sans-serif;letter-spacing:.05em;text-transform:uppercase;color:' + muted + '">' + t + '</th>';
  const thn = t => th(t).replace('text-align:left', 'text-align:right');
  const td = t => '<td style="padding:6px 8px;border-bottom:1px solid ' + line + ';font:14px Arial,sans-serif;color:' + ink + '">' + t + '</td>';
  const tdn = t => td(t).replace('font:14px', 'text-align:right;font:13px Consolas,Menlo,monospace;font-size:13px');

  let html = '<div style="max-width:640px;font-family:Arial,sans-serif;color:' + ink + '">';
  const good = '#23794a';
  const h3 = t => '<h3 style="font:600 13px Arial,sans-serif;letter-spacing:.06em;text-transform:uppercase;color:' + muted + ';margin:18px 0 6px">' + t + '</h3>';
  const h2 = t => '<h2 style="font:700 16px Arial,sans-serif;color:' + ink + ';margin:22px 0 4px;padding-top:12px;border-top:2px solid ' + ink + '">' + t + '</h2>';
  const tbl = '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%"><tr>';
  const pctCell = v => v == null ? '—' : '<span style="font-weight:700;color:' + (v * 100 >= res.goal ? good : bad) + '">' + pct_(v) + '</span>';
  html += '<div style="background:#132a38;color:#e9f1f6;padding:14px 18px;font:700 18px Arial,sans-serif;letter-spacing:.03em;text-transform:uppercase">' + (res ? 'Print room daily report' : 'Print room forecast') + '</div>';
  html += '<p style="font:15px Arial,sans-serif;margin:16px 0 4px"><b>' + esc_(label) + '</b>' + (sentBy ? ' · sent by ' + esc_(sentBy) : '') + '</p>';
  if (note) html += '<p style="font:14px Arial,sans-serif;margin:8px 0;padding:10px 12px;background:#eef2f5;border-radius:4px">' + esc_(note).replace(/\n/g, '<br>') + '</p>';
  if (res) {
    html += h2('Yesterday · ' + esc_(res.label));
    const ok = res.productivity != null && res.productivity * 100 >= res.goal;
    html += '<table role="presentation" cellspacing="0" cellpadding="0" style="margin:10px 0;border-collapse:collapse"><tr>' +
      tile(res.productivity == null ? '—' : '<span style="color:' + (ok ? good : bad) + '">' + pct_(res.productivity) + '</span>', 'Productivity', 'goal ' + res.goal + '%' + (res.productivity == null ? ' · no hours worked entered' : ok ? ' · met' : ' · below goal')) +
      tile(hrs_(res.earned), 'Earned hours', 'work completed, at standard times') +
      tile(res.worked ? hrs_(res.worked) : '—', 'Hours worked', res.worked ? hrs_(res.logged) + ' logged on timers' : '') + '</tr></table>';
    if (res.worked && res.worked - res.logged >= 0.25) html += '<p style="font:13px Arial,sans-serif;color:' + muted + ';margin:4px 0">' + hrs_(res.worked - res.logged) + ' of the hours worked weren\'t logged on any task.</p>';
    if (res.techs.length) {
      html += h3('By tech') + tbl + th('Tech') + thn('Worked') + thn('Logged') + thn('Earned') + thn('Productivity') + '</tr>';
      res.techs.forEach(t => { html += '<tr>' + td(esc_(t.name)) + tdn(t.worked ? hrs_(t.worked) : '—') + tdn(hrs_(t.logged)) + tdn(hrs_(t.earned)) + tdn(pctCell(t.worked ? t.earned / t.worked : null)) + '</tr>'; });
      html += '</table>';
    }
    if (res.products.length) {
      const anyPlan = res.products.some(p => p.planned != null);
      html += h3('Finished') + tbl + th('Product') + (anyPlan ? thn('Planned') : '') + thn('Finished') + '</tr>';
      res.products.forEach(p => { html += '<tr>' + td(esc_(p.name)) + (anyPlan ? tdn(p.planned == null ? '—' : num_(p.planned)) : '') + tdn('<b>' + num_(p.finished) + '</b>') + '</tr>'; });
      html += '</table>';
    }
    if (res.steps.length) {
      html += h3('Where the time went') + tbl + th('Product') + th('Step') + thn('Units') + thn('Earned hrs') + '</tr>';
      res.steps.forEach(x => { html += '<tr>' + td(esc_(x.product)) + td(esc_(x.step)) + tdn(x.units == null ? '—' : num_(x.units)) + tdn(hrs_(x.earned)) + '</tr>'; });
      html += '</table>';
    }
    if (res.unrated.length) html += '<p style="font:13px Arial,sans-serif;color:' + muted + ';margin:8px 0">No standard time yet, so not counted in earned hours: ' + res.unrated.map(u => esc_(u.product) + ' · ' + esc_(u.step) + (u.units ? ' (' + num_(u.units) + ')' : '')).join(', ') + '.</p>';
    html += h2("Today's forecast · " + esc_(label));
  }
  html += '<table role="presentation" cellspacing="0" cellpadding="0" style="margin:14px 0;border-collapse:collapse"><tr>' +
    tile(hrs_(labor), 'Labor hours', 'incl. overhead') +
    tile(num_(people, 1), 'People needed', hpp ? 'at ' + num_(hpp, 2).replace(/\.?0+$/, '') + ' productive hrs each' : '') +
    tile(printer ? hrs_(printer) : '—', 'Printer hours', cap ? 'of ' + hrs_(cap) + ' available' : '') + '</tr></table>';
  if (cap && printer > cap) html += '<p style="font:14px Arial,sans-serif;color:' + bad + '">Printers are over capacity for this volume. The print queue, not staffing, sets the pace.</p>';
  if (missing) html += '<p style="font:13px Arial,sans-serif;color:' + muted + '">' + missing + ' step' + (missing > 1 ? 's have' : ' has') + ' no target yet and ' + (missing > 1 ? 'are' : 'is') + ' left out of these hours.</p>';
  if (behind) html += '<p style="font:14px Arial,sans-serif;margin:6px 0">Furthest behind: <b>' + esc_(behind.name) + '</b>, oldest order ' + esc_(behind.oldest) + ' (' + daysTxt(behind.days) + ' waiting).</p>';
  if (products.length) {
    const anyOld = products.some(p => p.oldest);
    html += '<h3 style="font:600 13px Arial,sans-serif;letter-spacing:.06em;text-transform:uppercase;color:' + muted + ';margin:18px 0 6px">Planned units' + (anyOld ? ' and backlog' : '') + '</h3><table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%"><tr>' + th('Product') + thn('Units') + (anyOld ? th('Oldest order') + thn('Waiting') : '') + '</tr>';
    products.forEach(p => { html += '<tr>' + td(esc_(p.name)) + tdn(p.units ? num_(p.units) : '—') + (anyOld ? td(p.oldest ? esc_(p.oldest) : '—') + tdn(p.days == null ? '—' : '<span style="font-weight:700;color:' + (p.days >= 3 ? bad : p.days >= 2 ? '#a86300' : '#23794a') + '">' + daysTxt(p.days) + '</span>') : '') + '</tr>'; });
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
  if (res) {
    text += 'YESTERDAY · ' + res.label + '\nProductivity: ' + pct_(res.productivity) + ' (goal ' + res.goal + '%)\nEarned hours: ' + hrs_(res.earned) + '\nHours worked: ' + (res.worked ? hrs_(res.worked) : '—') + ' (' + hrs_(res.logged) + ' logged)\n';
    if (res.techs.length) text += res.techs.map(t => '  ' + t.name + ': worked ' + (t.worked ? hrs_(t.worked) : '—') + ', earned ' + hrs_(t.earned) + ', ' + pct_(t.worked ? t.earned / t.worked : null)).join('\n') + '\n';
    if (res.products.length) text += 'Finished\n' + res.products.map(p => '  ' + p.name + ': ' + num_(p.finished) + (p.planned != null ? ' of ' + num_(p.planned) + ' planned' : '')).join('\n') + '\n';
    text += '\nTODAY · ' + label + '\n';
  }
  text += 'Labor hours: ' + hrs_(labor) + ' (incl. overhead)\nPeople needed: ' + num_(people, 1) + '\nPrinter hours: ' + (printer ? hrs_(printer) : '—') + (cap ? ' of ' + hrs_(cap) + ' available' : '') + '\n\n';
  if (behind) text += 'Furthest behind: ' + behind.name + ', oldest order ' + behind.oldest + ' (' + daysTxt(behind.days) + ' waiting)\n\n';
  if (products.length) text += 'Planned units\n' + products.map(p => '  ' + p.name + ': ' + (p.units ? num_(p.units) : '—') + (p.oldest ? '  · oldest order ' + p.oldest + ' (' + daysTxt(p.days) + ')' : '')).join('\n') + '\n\n';
  if (steps.length) text += 'Hours by step\n' + steps.map(x => '  ' + x.product + ' · ' + x.step + ': ' + hrs_(x.hours) + ' h').join('\n') + '\n';

  MailApp.sendEmail({ to: list.join(','), subject: subject, htmlBody: html, body: text, name: 'Print Room Tracker' });
  props.setProperty(dayKey, String(sentToday + 1));

  // keep a record of what went out
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let log = ss.getSheetByName('Reports');
  if (!log) { log = ss.insertSheet('Reports'); log.setFrozenRows(1); }
  if (log.getLastColumn() < REPORT_COLS.length) log.getRange(1, 1, 1, REPORT_COLS.length).setValues([REPORT_COLS]).setFontWeight('bold');
  const reportDate = /^\d{4}-\d{2}-\d{2}$/.test(String(r.date || '')) ? String(r.date) : '';
  const plan = {};
  if (r.plan && typeof r.plan === 'object') Object.keys(r.plan).slice(0, 40).forEach(k => { const n = Number(r.plan[k]); if (k.length <= 40 && isFinite(n) && n > 0) plan[k] = n; });
  log.appendRow([new Date(), label, period, Math.round(labor * 10) / 10, Math.round(people * 10) / 10, Math.round(printer * 10) / 10, sentBy, list.join(', '),
    reportDate, JSON.stringify(plan), res ? JSON.stringify({ date: res.label, productivity: res.productivity, earned: res.earned, worked: res.worked, logged: res.logged, techs: res.techs }) : '']);
  return { ok: true, sentTo: list.length };
}

const REPORT_COLS = ['sentAt','report','period','laborHours','peopleNeeded','printerHours','sentBy','recipients','reportDate','plan','yesterday'];

function parseResults_(x) {
  if (!x || typeof x !== 'object') return null;
  const n = v => { const k = Number(v); return isFinite(k) && k >= 0 ? k : 0; };
  const goal = Math.min(200, Math.max(1, Number(x.goal) || 95));
  const techs = (Array.isArray(x.techs) ? x.techs : []).slice(0, 30).map(t => ({ name: str_(t.name, 60), worked: n(t.worked), earned: n(t.earned), logged: n(t.logged) })).filter(t => t.name && (t.worked || t.earned || t.logged));
  const products = (Array.isArray(x.products) ? x.products : []).slice(0, 30).map(p => ({ name: str_(p.name, 60), finished: n(p.finished), planned: p.planned == null || p.planned === '' ? null : n(p.planned) })).filter(p => p.name && (p.finished || p.planned));
  const steps = (Array.isArray(x.steps) ? x.steps : []).slice(0, 12).map(s => ({ product: str_(s.product, 60), step: str_(s.step, 80), units: s.units == null ? null : n(s.units), earned: n(s.earned) })).filter(s => s.earned > 0);
  const unrated = (Array.isArray(x.unrated) ? x.unrated : []).slice(0, 12).map(s => ({ product: str_(s.product, 60), step: str_(s.step, 80), units: n(s.units) }));
  const worked = n(x.worked), earned = n(x.earned), logged = n(x.logged);
  return { label: str_(x.label, 60), goal: goal, worked: worked, earned: earned, logged: logged, productivity: worked > 0 ? earned / worked : null, techs: techs, products: products, steps: steps, unrated: unrated };
}

// The planned units from the last day report sent for a given date, so the next day's report can compare plan to actual.
function getPlan_(date) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { ok: true, plan: null };
  const log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reports');
  if (!log || log.getLastRow() < 2 || log.getLastColumn() < REPORT_COLS.length) return { ok: true, plan: null };
  const rows = log.getRange(2, 1, log.getLastRow() - 1, REPORT_COLS.length).getValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    const d = rows[i][8] instanceof Date ? Utilities.formatDate(rows[i][8], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(rows[i][8]);
    if (rows[i][2] === 'day' && d === date) { try { return { ok: true, plan: JSON.parse(rows[i][9] || '{}'), sentAt: rows[i][0] instanceof Date ? rows[i][0].getTime() : null }; } catch (e) { return { ok: true, plan: null }; } }
  }
  return { ok: true, plan: null };
}

/** Run once from the editor after pasting this version, to allow the script to send email as you. */
function authorizeEmail() { Logger.log('Emails left today: ' + MailApp.getRemainingDailyQuota()); }

/** Run once from the editor (select setup, press Run) to create the tabs and grant access. */
function setup() { entriesSheet_(); configSheet_(); }
