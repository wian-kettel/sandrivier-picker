/**
 * Sandrivier picker tracker: Google Sheet backend (v3, Oct 2026)
 *
 * The farm phone (scanning screen on GitHub Pages) sends scans here in batches.
 * Everything is stored in this Google Sheet; nothing is stored on GitHub.
 *
 * Tabs (created or extended by setup()):
 *   Pickers  : Picker no | Name | Tag UID | Active | Registered | Type (Day/Contract) | Old tags
 *   Taps     : Received | Date | Time | Picker no | Name | Tag UID | Block | Status |
 *              Team | Device | Method | Scan ID | Scan time | Bags
 *   Bins     : Date | Team | Export bins (plastic) | Juice bins (wooden) | Updated | Bags scanned
 *   Settings : see setup()
 *   Load test: results of loadTest()
 *
 * Status: OK (counted), DUPLICATE (same picker inside the repeat window, not counted),
 *         UNKNOWN (band not registered, not counted).
 * Method: NFC (band tapped), MANUAL (number typed on the phone), PAPER (day typed in from the clipboard),
 *         LINK (iPhone link test).
 * Bags:   1 for every scan; a PAPER row carries the picker's total for that day.
 */

const TZ = 'Africa/Johannesburg';
const DEMO_PREFIX = 'D0:00:00';
const TAP_HEADERS = ['Received', 'Date', 'Time', 'Picker no', 'Name', 'Tag UID', 'Block', 'Status',
                     'Team', 'Device', 'Method', 'Scan ID', 'Scan time', 'Bags'];
const PICKER_HEADERS = ['Picker no', 'Name', 'Tag UID', 'Active', 'Registered', 'Type', 'Old tags'];
const BIN_HEADERS = ['Date', 'Team', 'Export bins (plastic)', 'Juice bins (wooden)', 'Updated', 'Bags scanned'];
const TEAMS = ['Team 1', 'Team 2'];
const BIN_TEAMS = ['All teams', 'Team 1', 'Team 2'];

/* ---------------- setup ---------------- */

function setup() {
  pickers_(); taps_(); bins_();
  const s = settings_();
  const have = s.getDataRange().getValues().map(r => r[0]);
  const defaults = [
    ['Current block', 'Blok 1'],
    ['Blocks', 'Blok 1,Blok 2,Blok 3,Blok 4,Blok 5,Blok 6,Blok 7,Blok 8'],
    ['Farm PIN', String(Math.floor(1000 + Math.random() * 9000))],
    ['Repeat window minutes', 2],
    ['Bags per bin', 44],
    ['Export bin target', 72],
    ['Expected export %', 50],
    ['Picking sessions', '06:30-08:00,09:00-12:00,13:00-17:00']
  ];
  defaults.forEach(d => { if (have.indexOf(d[0]) < 0) s.appendRow(d); });
}

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.setFrozenRows(1); }
  const cur = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  if (cur.join('|') !== headers.join('|')) {
    headers.forEach((h, i) => { if (!cur[i]) sh.getRange(1, i + 1).setValue(h); });
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}
function pickers_()  { return sheet_('Pickers', PICKER_HEADERS); }
function taps_()     { return sheet_('Taps', TAP_HEADERS); }
function bins_()     { return sheet_('Bins', BIN_HEADERS); }
function settings_() { return sheet_('Settings', ['Setting', 'Value']); }

/* ---------------- helpers ---------------- */

function normUid_(raw) {
  const hex = String(raw || '').toUpperCase().replace(/[^0-9A-F]/g, '');
  if (hex.length < 4) return '';
  return hex.match(/.{1,2}/g).join(':');
}
function day_(d)  { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function time_(d) { return Utilities.formatDate(d, TZ, 'HH:mm'); }
function dstr_(v) { return v instanceof Date ? day_(v) : String(v || ''); }
function num_(v)  { const n = parseFloat(String(v).replace(',', '.')); return isNaN(n) ? null : n; }
function esc_(s)  { return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function addDays_(date, n) { return day_(new Date(new Date(date + 'T12:00:00+02:00').getTime() + n * 86400000)); }
function at_(date, hhmm) { return new Date(date + 'T' + hhmm + ':00+02:00'); }

function getSetting_(key) {
  const v = settings_().getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === key) return String(v[i][1]);
  return '';
}
function setSetting_(key, val) {
  const sh = settings_();
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === key) { sh.getRange(i + 1, 2).setValue(val); return; }
  sh.appendRow([key, val]);
}
function repeatMs_() { return (num_(getSetting_('Repeat window minutes')) || 2) * 60000; }
function blocks_()   { return getSetting_('Blocks').split(',').map(s => s.trim()).filter(String); }
function sessions_() {
  return getSetting_('Picking sessions').split(',').map(s => s.trim().split('-')).filter(p => p.length === 2)
    .map(p => ({ start: p[0].trim(), end: p[1].trim() }));
}
// picking minutes between two moments on one day, breaks excluded
function pickingMinutes_(date, from, to) {
  let m = 0;
  sessions_().forEach(s => {
    const a = Math.max(at_(date, s.start).getTime(), from.getTime());
    const b = Math.min(at_(date, s.end).getTime(), to.getTime());
    if (b > a) m += (b - a) / 60000;
  });
  return m;
}

function checkPin_(pin) {
  const want = getSetting_('Farm PIN');
  if (!want || String(pin || '').trim() !== want) throw new Error('Wrong farm PIN.');
}

/* ---------------- pickers ---------------- */

function pickerData_() {
  const v = pickers_().getDataRange().getValues();
  const byUid = {}, byNo = {}, list = [];
  for (let i = 1; i < v.length; i++) {
    const no = String(v[i][0]).trim();
    if (!no) continue;
    const p = {
      row: i + 1, no: no, name: String(v[i][1]), uid: normUid_(v[i][2]),
      active: String(v[i][3]).toUpperCase() !== 'NO',
      type: String(v[i][5] || 'Day') === 'Contract' ? 'Contract' : 'Day',
      oldTags: String(v[i][6] || '')
    };
    if (!p.active) continue;
    list.push(p);
    byNo[no] = p;
    if (p.uid) byUid[p.uid] = p;
  }
  return { byUid: byUid, byNo: byNo, list: list };
}

/* ---------------- reading taps fast ---------------- */
// Script property "idx" remembers the first Taps row of each day, so the dashboard reads one day, not the season.

function idx_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('idx') || '{}'); } catch (e) { return {}; } }
function saveIdx_(o) { PropertiesService.getScriptProperties().setProperty('idx', JSON.stringify(o)); }
function resetIdx_() { PropertiesService.getScriptProperties().deleteProperty('idx'); }

function firstRowFor_(date) {
  const idx = idx_();
  if (idx[date]) return idx[date];
  const sh = taps_();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const col = sh.getRange(2, 2, last - 1, 1).getValues();          // Date column only
  for (let i = 0; i < col.length; i++) {
    if (dstr_(col[i][0]) >= date) { idx[date] = i + 2; saveIdx_(idx); return i + 2; }
  }
  return null;
}

function readTapsFrom_(fromRow) {
  const sh = taps_();
  const last = sh.getLastRow();
  if (!fromRow || last < fromRow) return [];
  return sh.getRange(fromRow, 1, last - fromRow + 1, TAP_HEADERS.length).getValues().map(r => {
    const rec = r[0] instanceof Date ? r[0] : new Date(r[0]);
    const st = r[12] instanceof Date ? r[12] : null;
    const ts = st || rec;
    return {
      ts: ts, date: r[1] ? dstr_(r[1]) : (isNaN(ts.getTime()) ? '' : day_(ts)),
      no: String(r[3]).trim(), name: String(r[4]), uid: normUid_(r[5]),
      block: String(r[6]), status: String(r[7]), team: String(r[8] || ''),
      device: String(r[9] || ''), method: String(r[10] || 'NFC'), id: String(r[11] || ''),
      bags: Number(r[13]) || 1
    };
  }).filter(t => !isNaN(t.ts.getTime()));
}
function tapsForDate_(date) { return readTapsFrom_(firstRowFor_(date)).filter(t => t.date === date); }

/* ---------------- web entry points ---------------- */

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.api === 'roster') {
      checkPin_(p.pin);
      const pd = pickerData_();
      return json_({
        ok: true,
        pickers: pd.list.map(x => ({ no: x.no, name: x.name, uid: x.uid, type: x.type, active: true })),
        blocks: blocks_(),
        currentBlock: getSetting_('Current block'),
        repeatMinutes: num_(getSetting_('Repeat window minutes')) || 2,
        teams: TEAMS
      });
    }
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
  if (p.tag) return page_(tapHtml_(normUid_(p.tag)), 'Bag scan');
  return HtmlService.createTemplateFromFile('Dashboard').evaluate()
    .setTitle('Sandrivier picking')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    checkPin_(body.pin);
    if (body.action === 'scans') return json_(ingestScans_(body.scans || [], body.device || '', body.team || ''));
    if (body.action === 'register') return json_({ ok: true, picker: registerPicker(body.uid, body.no, body.name, body.type, body.replace) });
    return json_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

function page_(html, title) {
  return HtmlService.createHtmlOutput(html).setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------------- scans from the farm phone ---------------- */

// scans: [{id, uid, no, block, t (ms epoch), method, status}]
function ingestScans_(scans, device, team) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const pd = pickerData_();
    const now = new Date();
    const today = day_(now);
    let minDate = today;
    scans.forEach(s => { if (s && s.t) { const d = day_(new Date(Number(s.t))); if (d < minDate) minDate = d; } });
    // read only from the first row of the oldest day in this batch (yesterday at most in practice)
    const from = firstRowFor_(minDate) || firstRowFor_(today);
    const sh = taps_();
    const seen = {}, lastOk = {};
    if (from) {
      const last = sh.getLastRow();
      if (last >= from) {
        sh.getRange(from, 4, last - from + 1, 10).getValues().forEach(r => {   // cols D..M
          const id = String(r[8] || ''); if (id) seen[id] = true;
          if (String(r[4]) === 'OK' && String(r[7]) !== 'PAPER') {
            const st = r[9] instanceof Date ? r[9] : null;
            const no = String(r[0]).trim();
            if (st && (!lastOk[no] || st > lastOk[no])) lastOk[no] = st;
          }
        });
      }
    }
    const windowMs = repeatMs_();
    const rows = [], accepted = [];
    scans.slice().sort((a, b) => a.t - b.t).forEach(s => {
      if (!s || !s.id) return;
      accepted.push(s.id);
      if (seen[s.id]) return;                    // already stored (phone re-sent after a dropped reply)
      seen[s.id] = true;
      let uid = normUid_(s.uid);
      let p = uid ? pd.byUid[uid] : null;
      if (!p && s.no && pd.byNo[String(s.no)]) { p = pd.byNo[String(s.no)]; uid = p.uid; }
      const st = new Date(Number(s.t) || now.getTime());
      let status;
      if (!p) status = 'UNKNOWN';
      else if (s.status === 'DUPLICATE') status = 'DUPLICATE';
      else if (lastOk[p.no] && st - lastOk[p.no] < windowMs) status = 'DUPLICATE';
      else status = 'OK';
      if (status === 'OK') lastOk[p.no] = st;
      rows.push([now, day_(st), Utilities.formatDate(st, TZ, 'HH:mm:ss'),
                 p ? p.no : String(s.no || ''), p ? p.name : '', uid, String(s.block || ''), status,
                 String(team || ''), String(device || ''), String(s.method || 'NFC'), String(s.id), st, 1]);
    });
    appendTaps_(rows);
    return { ok: true, accepted: accepted, stored: rows.length };
  } finally {
    lock.releaseLock();
  }
}

function appendTaps_(rows) {
  if (!rows.length) return;
  const sh = taps_();
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, rows.length, TAP_HEADERS.length).setValues(rows);
  const idx = idx_(); let changed = false;
  rows.forEach((r, i) => { const d = r[1]; if (!idx[d]) { idx[d] = start + i; changed = true; } });
  if (changed) saveIdx_(idx);
}

/* ---------------- iPhone link scans (prototype, kept for testing) ---------------- */

function recordTap_(uid) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const p = pickerData_().byUid[uid];
    if (!p) return { state: 'unknown' };
    const now = new Date();
    const today = day_(now);
    let lastOk = null, count = 0;
    tapsForDate_(today).forEach(t => {
      if (t.no !== p.no || t.status !== 'OK') return;
      count += t.bags;
      if (t.method !== 'PAPER' && (!lastOk || t.ts > lastOk)) lastOk = t.ts;
    });
    const windowMs = repeatMs_();
    const ms = lastOk ? now - lastOk : Infinity;
    const status = ms < windowMs ? 'DUPLICATE' : 'OK';
    const block = getSetting_('Current block');
    appendTaps_([[now, today, Utilities.formatDate(now, TZ, 'HH:mm:ss'), p.no, p.name, uid, block, status,
                  '', 'iPhone link', 'LINK', Utilities.getUuid(), now, 1]]);
    if (status === 'OK') count++;
    return { state: status === 'OK' ? 'ok' : 'duplicate', picker: p, count: count, block: block,
             wait: Math.max(1, Math.ceil((windowMs - ms) / 1000)) };
  } finally {
    lock.releaseLock();
  }
}

function tapHtml_(uid) {
  const css = '<style>body{margin:0;font-family:-apple-system,Segoe UI,Roboto,sans-serif;}' +
    '.p{min-height:100vh;display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center;padding:24px;box-sizing:border-box;color:#fff}' +
    '.big{font-size:64px;font-weight:800;line-height:1}.name{font-size:30px;font-weight:700;margin:16px 0 4px}' +
    '.sub{font-size:18px;opacity:.9}</style>';
  if (!uid) return css + '<div class="p" style="background:#6b7280"><div class="name">No band number in link</div></div>';
  const r = recordTap_(uid);
  if (r.state === 'ok') {
    return css + '<div class="p" style="background:#15803d"><div class="sub">BAG COUNTED</div>' +
      '<div class="name">' + esc_(r.picker.name) + ' (#' + esc_(r.picker.no) + ')</div>' +
      '<div class="big">' + r.count + '</div><div class="sub">bags today · ' + esc_(r.block) + '</div></div>';
  }
  if (r.state === 'duplicate') {
    return css + '<div class="p" style="background:#b45309"><div class="sub">NOT COUNTED</div>' +
      '<div class="name">' + esc_(r.picker.name) + '</div>' +
      '<div class="sub">Already scanned a moment ago.<br>Wait ' + r.wait + ' seconds for the next bag.</div></div>';
  }
  return css + '<div class="p" style="background:#1d4ed8"><div class="name">Band not registered</div>' +
    '<div class="sub">Register it on the farm phone.</div></div>';
}

/* ---------------- called from pages or the phone ---------------- */

/**
 * Link a band to a picker.
 *  - number not on the list yet: new picker (name required)
 *  - number on the list without a band (pre-loaded): band linked, name optional
 *  - number already has a band: refused with HAS_BAND unless replace is true;
 *    with replace the old band is kept in "Old tags" and the picker's history stays with the picker number
 */
function registerPicker(uid, no, name, type, replace) {
  uid = normUid_(uid); no = String(no || '').trim(); name = String(name || '').trim();
  if (!uid) throw new Error('No band was read.');
  if (!no) throw new Error('Type the picker number.');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const pd = pickerData_();
    const sh = pickers_();
    if (pd.byUid[uid]) {
      const o = pd.byUid[uid];
      if (o.no === no) return pub_(o);
      throw new Error('This band is already linked to ' + o.name + ' (#' + o.no + ').');
    }
    const p = pd.byNo[no];
    if (p && !p.uid) {
      sh.getRange(p.row, 3).setValue(uid);
      if (name && !p.name) sh.getRange(p.row, 2).setValue(name);
      if (type) sh.getRange(p.row, 6).setValue(type === 'Contract' ? 'Contract' : 'Day');
      p.uid = uid; if (name && !p.name) p.name = name;
      return pub_(p);
    }
    if (p && p.uid) {
      if (!replace) throw new Error('HAS_BAND:' + p.name);
      sh.getRange(p.row, 7).setValue(p.oldTags ? p.oldTags + ', ' + p.uid : p.uid);
      sh.getRange(p.row, 3).setValue(uid);
      p.uid = uid;
      return pub_(p);
    }
    if (!name) throw new Error('Picker ' + no + ' is not on the list yet. Type the name too.');
    const row = sh.getLastRow() + 1;
    sh.getRange(row, 1).setNumberFormat('@');
    sh.getRange(row, 1, 1, PICKER_HEADERS.length).setValues([[no, name, uid, 'YES', new Date(), type === 'Contract' ? 'Contract' : 'Day', '']]);
    return { no: no, name: name, uid: uid, type: type === 'Contract' ? 'Contract' : 'Day', active: true };
  } finally {
    lock.releaseLock();
  }
}
function pub_(p) { return { no: p.no, name: p.name, uid: p.uid, type: p.type, active: true }; }

function setBlock(block) {
  if (blocks_().indexOf(block) < 0) throw new Error('Unknown block: ' + block);
  setSetting_('Current block', block);
  return block;
}

// Fanus's actual bin count for a day (as on the whiteboard). Also stores the bags scanned that day,
// which is what lets the app learn the real bags per bin.
function saveBins(date, team, exportBins, juiceBins) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date || '')) throw new Error('Pick a day first.');
  if (BIN_TEAMS.indexOf(team) < 0) throw new Error('Pick a team.');
  const e = Math.max(0, Math.round(num_(exportBins) || 0));
  const j = Math.max(0, Math.round(num_(juiceBins) || 0));
  const bags = tapsForDate_(date).filter(t => t.status === 'OK' && (team === 'All teams' || t.team === team))
    .reduce((s, t) => s + t.bags, 0);
  const sh = bins_();
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (dstr_(v[i][0]) === date && String(v[i][1]) === team) {
      sh.getRange(i + 1, 3, 1, 4).setValues([[e, j, new Date(), bags]]);
      return true;
    }
  }
  const row = sh.getLastRow() + 1;
  sh.getRange(row, 1).setNumberFormat('@');
  sh.getRange(row, 1, 1, BIN_HEADERS.length).setValues([[date, team, e, j, new Date(), bags]]);
  return true;
}

// A day typed in from the clipboard: one line per picker, "number bags" (e.g. "12 148").
function saveBackfill(date, team, block, text) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date || '')) throw new Error('Pick a day first.');
  const pd = pickerData_();
  const lines = String(text || '').split(/\r?\n/).map(l => l.trim()).filter(String);
  const rows = [], problems = [];
  const now = new Date();
  lines.forEach((l, i) => {
    const m = l.match(/^([^\s,;:]+)[\s,;:]+(\d+)$/);
    if (!m) { problems.push('Line ' + (i + 1) + ': "' + l + '" is not "number bags"'); return; }
    const p = pd.byNo[m[1]];
    if (!p) { problems.push('Line ' + (i + 1) + ': picker ' + m[1] + ' is not on the list'); return; }
    rows.push([now, date, '12:00:00', p.no, p.name, p.uid, String(block || ''), 'OK', String(team || ''),
               'paper', 'PAPER', Utilities.getUuid(), at_(date, '12:00'), Number(m[2])]);
  });
  if (problems.length) throw new Error(problems.join('\n'));
  if (!rows.length) throw new Error('Nothing to save.');
  const already = tapsForDate_(date).filter(t => t.method === 'PAPER').length;
  appendTaps_(rows);
  return { saved: rows.length, bags: rows.reduce((s, r) => s + r[13], 0), alreadyHadPaper: already };
}

/* ---------------- dashboard ---------------- */

function binActuals_() {
  const out = {};
  bins_().getDataRange().getValues().slice(1).forEach(r => {
    const d = dstr_(r[0]); if (!d) return;
    const o = out[d] || (out[d] = { teams: {} });
    o.teams[String(r[1])] = { exp: Number(r[2]) || 0, juice: Number(r[3]) || 0, bags: Number(r[5]) || 0 };
  });
  Object.keys(out).forEach(d => {
    const t = out[d].teams;
    const src = t['All teams'] ? [t['All teams']] : Object.keys(t).map(k => t[k]);
    out[d].exp = src.reduce((s, x) => s + x.exp, 0);
    out[d].juice = src.reduce((s, x) => s + x.juice, 0);
    out[d].bags = src.reduce((s, x) => s + x.bags, 0);
  });
  return out;
}

// bags per bin and export share learned from the last 7 days Fanus entered (falls back to Settings)
function calibration_(beforeDate, actuals) {
  const days = Object.keys(actuals).filter(d => d < beforeDate).sort().slice(-7);
  let bags = 0, bins = 0, exp = 0, binsAll = 0, nb = 0, ne = 0;
  days.forEach(d => {
    const a = actuals[d]; const b = a.exp + a.juice;
    if (b > 0) { exp += a.exp; binsAll += b; ne++; }
    if (b > 0 && a.bags > 0) { bags += a.bags; bins += b; nb++; }
  });
  const setBpb = num_(getSetting_('Bags per bin')) || 44;
  const setExp = (num_(getSetting_('Expected export %')) || 50) / 100;
  return {
    bagsPerBin: nb ? bags / bins : setBpb, bpbFrom: nb ? ('last ' + nb + ' day' + (nb > 1 ? 's' : '')) : 'setting',
    exportShare: ne ? exp / binsAll : setExp, expFrom: ne ? ('last ' + ne + ' day' + (ne > 1 ? 's' : '')) : 'setting'
  };
}

function dashboardData(dateStr, rankBy) {
  const now = new Date();
  const today = day_(now);
  const date = /^\d{4}-\d{2}-\d{2}$/.test(dateStr || '') ? dateStr : today;
  rankBy = rankBy === 'bags' ? 'bags' : 'hour';
  const pd = pickerData_();
  const all = tapsForDate_(date);
  const ok = all.filter(t => t.status === 'OK');
  const bags = ok.reduce((s, t) => s + t.bags, 0);
  const actuals = binActuals_();
  const cal = calibration_(date, actuals);
  const target = num_(getSetting_('Export bin target')) || 72;
  const sess = sessions_();
  const dayEnd = sess.length ? at_(date, sess[sess.length - 1].end) : at_(date, '17:00');

  const rows = perPicker_(date, ok, pd);
  const ranked = rankBy === 'hour'
    ? rows.filter(r => r.perHour !== null).sort((a, b) => b.perHour - a.perHour || b.bags - a.bags)
    : rows.slice();

  // estimate, run rate, forecast
  const estBins = bags / cal.bagsPerBin;
  const est = { bins: round1_(estBins), exp: round1_(estBins * cal.exportShare), juice: round1_(estBins * (1 - cal.exportShare)) };
  let forecast = null;
  if (date === today && ok.length) {
    const scanned = ok.filter(t => t.method !== 'PAPER');
    const firstTs = scanned.reduce((m, t) => (!m || t.ts < m) ? t.ts : m, null);
    const elapsed = firstTs ? pickingMinutes_(date, firstTs, now) : 0;
    // rate over the last 60 picking minutes (or since the first scan if the day is young)
    let windowStart = now, mins = 0;
    while (mins < 60 && windowStart > (firstTs || now)) { windowStart = new Date(windowStart.getTime() - 5 * 60000); mins = pickingMinutes_(date, windowStart, now); }
    const recentBags = scanned.filter(t => t.ts >= windowStart).reduce((s, t) => s + t.bags, 0);
    const rateMins = Math.max(mins, 1);
    const rate = elapsed >= 15 ? recentBags / (rateMins / 60) : null;      // bags per picking hour
    const remaining = now < dayEnd ? pickingMinutes_(date, now, dayEnd) : 0;
    if (rate !== null) {
      const endBags = bags + rate * remaining / 60;
      const endBins = endBags / cal.bagsPerBin;
      const bagsForTarget = target / cal.exportShare * cal.bagsPerBin;
      let reachAt = null;
      if (bags >= bagsForTarget) reachAt = 'reached';
      else if (rate > 0) {
        let need = (bagsForTarget - bags) / rate * 60;           // picking minutes still needed
        for (let i = 0; i < sess.length && reachAt === null; i++) {
          const a = Math.max(at_(date, sess[i].start).getTime(), now.getTime());
          const b = at_(date, sess[i].end).getTime();
          if (b <= a) continue;
          const len = (b - a) / 60000;
          if (need <= len) reachAt = time_(new Date(a + need * 60000)); else need -= len;
        }
      }
      forecast = {
        ratePerHour: Math.round(rate), binsPerHour: round1_(rate / cal.bagsPerBin),
        endBins: round1_(endBins), endExp: round1_(endBins * cal.exportShare), endTime: time_(dayEnd),
        reachAt: reachAt, remainingMin: Math.round(remaining),
        neededPerHour: remaining > 0 ? Math.max(0, Math.round((bagsForTarget - bags) / (remaining / 60))) : null
      };
    }
  }

  const act = actuals[date] || null;
  const problems = [];
  rows.filter(r => r.manual >= 3).forEach(r => problems.push(r.name + ' (#' + r.no + ') had ' + r.manual + ' bags typed in: band may be broken or forgotten.'));
  const unknown = all.filter(t => t.status === 'UNKNOWN').length;
  if (unknown) problems.push(unknown + ' scan' + (unknown > 1 ? 's' : '') + ' of a band that is not registered.');
  const dups = all.filter(t => t.status === 'DUPLICATE').length;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    date: date, isToday: date === today, rankBy: rankBy, updated: time_(now),
    target: target, bags: bags, pickers: rows.length, registered: pd.list.length,
    est: est, cal: { bagsPerBin: round1_(cal.bagsPerBin), bpbFrom: cal.bpbFrom, exportPct: Math.round(cal.exportShare * 100), expFrom: cal.expFrom },
    forecast: forecast,
    actual: act ? { exp: act.exp, juice: act.juice, total: act.exp + act.juice, exportPct: (act.exp + act.juice) ? Math.round(act.exp / (act.exp + act.juice) * 100) : null } : null,
    top: ranked.slice(0, 10), bottom: ranked.slice().reverse().slice(0, 10), rankedCount: ranked.length,
    problems: problems, duplicates: dups, manual: ok.filter(t => t.method === 'MANUAL').length,
    paper: ok.filter(t => t.method === 'PAPER').reduce((s, t) => s + t.bags, 0),
    teams: TEAMS, binTeams: BIN_TEAMS, blocks: blocks_(),
    sheetUrl: ss.getUrl(),
    xlsxUrl: 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx'
  };
}
// one row per picker for a day; per hour counts picking time only (breaks excluded), needs 1 hour of scans
function perPicker_(date, okTaps, pd) {
  const per = {};
  okTaps.forEach(t => {
    const r = per[t.no] || (per[t.no] = { no: t.no, name: t.name, bags: 0, manual: 0, paper: 0, first: null, last: null, blocks: {}, team: t.team });
    r.bags += t.bags;
    if (t.method === 'MANUAL') r.manual++;
    if (t.method === 'PAPER') { r.paper += t.bags; return; }
    if (!r.first || t.ts < r.first) r.first = t.ts;
    if (!r.last || t.ts > r.last) r.last = t.ts;
    r.blocks[t.block] = (r.blocks[t.block] || 0) + 1;
  });
  return Object.keys(per).map(k => {
    const r = per[k];
    const pm = (r.first && r.last) ? pickingMinutes_(date, r.first, r.last) : 0;
    return { no: r.no, name: r.name, bags: r.bags, manual: r.manual, paper: r.paper,
      perHour: pm >= 60 ? Math.round((r.bags - r.paper) / (pm / 60) * 10) / 10 : null,
      first: r.first ? time_(r.first) : '', last: r.last ? time_(r.last) : '', team: r.team || '',
      type: (pd.byNo[r.no] && pd.byNo[r.no].type) || 'Day',
      block: Object.keys(r.blocks).sort((a, b) => r.blocks[b] - r.blocks[a])[0] || '' };
  }).sort((a, b) => b.bags - a.bags || a.name.localeCompare(b.name));
}

function round1_(x) { return Math.round(x * 10) / 10; }

// The "Details" page: every picker, pairs, last 7 days, blocks, hours. Read on demand only.
function detailsData(dateStr) {
  const today = day_(new Date());
  const date = /^\d{4}-\d{2}-\d{2}$/.test(dateStr || '') ? dateStr : today;
  const pd = pickerData_();
  const from = addDays_(date, -6);
  const week = readTapsFrom_(firstRowFor_(from)).filter(t => t.date >= from && t.date <= date && t.status === 'OK');
  const wk = {};
  week.forEach(t => {
    const r = wk[t.no] || (wk[t.no] = { no: t.no, name: t.name, bags: 0, days: {} });
    r.bags += t.bags; r.days[t.date] = true;
  });
  const weekTop = Object.keys(wk).map(k => ({ no: wk[k].no, name: wk[k].name, bags: wk[k].bags, days: Object.keys(wk[k].days).length,
    type: (pd.byNo[k] && pd.byNo[k].type) || 'Day' })).sort((a, b) => b.bags - a.bags).slice(0, 10);
  const dayOk = week.filter(t => t.date === date);
  const byBlock = {}, hourly = {};
  dayOk.forEach(t => {
    byBlock[t.block || '(none)'] = (byBlock[t.block || '(none)'] || 0) + t.bags;
    if (t.method !== 'PAPER') { const h = Utilities.formatDate(t.ts, TZ, 'HH') + ':00'; hourly[h] = (hourly[h] || 0) + t.bags; }
  });
  const everyone = perPicker_(date, dayOk, pd);
  const base = everyone.filter(r => r.perHour !== null).sort((a, b) => b.perHour - a.perHour);
  const pairs = [];
  for (let i = 0, j = base.length - 1; i < j; i++, j--) pairs.push({ strong: base[i], weak: base[j] });
  return { date: date, rows: everyone, pairs: pairs, week: weekTop, byBlock: byBlock, hourly: hourly,
           teams: TEAMS, binTeams: BIN_TEAMS, blocks: blocks_() };
}

/* ---------------- demo data and load test (run from the editor) ---------------- */

function demoPickers_(n) {
  const pd = pickerData_();
  const names = ['Anna', 'Piet', 'Sanna', 'Koos', 'Martha', 'Jan', 'Lena', 'Dawid', 'Rosa', 'Hendrik',
                 'Elsie', 'Sakkie', 'Grieta', 'Willem', 'Bettie', 'Frans', 'Mina', 'Klaas', 'Sophie', 'Gert',
                 'Maria', 'Jacob', 'Lettie', 'Abel', 'Hester'];
  const out = [], add = [];
  for (let i = 0; i < n; i++) {
    const uid = DEMO_PREFIX + ':' + ('0' + ((i + 1) % 256).toString(16).toUpperCase()).slice(-2) + ':' + ('0' + Math.floor((i + 1) / 256)).slice(-2) + ':00';
    const no = String(900 + i + 1);
    const p = { no: no, name: 'Demo ' + names[i % names.length] + (i >= names.length ? ' ' + (Math.floor(i / names.length) + 1) : ''), uid: normUid_(uid), team: TEAMS[i % 2] };
    out.push(p);
    if (!pd.byNo[no]) add.push([no, p.name, p.uid, 'YES', new Date(), i % 3 === 0 ? 'Contract' : 'Day', '']);
  }
  if (add.length) { const sh = pickers_(); sh.getRange(sh.getLastRow() + 1, 1, add.length, PICKER_HEADERS.length).setValues(add); }
  return out;
}

// A normal demo day: 24 pickers, scans from 06:30 until now (or a full day).
function seedDemo() {
  setup();
  const today = day_(new Date());
  const ps = demoPickers_(24);
  const end = Math.min(Date.now(), at_(today, '17:00').getTime());
  const scans = [];
  ps.forEach((p, i) => {
    const every = (6 + (i * 7) % 19) * 60000;
    for (let t = at_(today, '06:30').getTime() + i * 60000; t < end; t += every) {
      if (pickingMinutes_(today, new Date(t - 1), new Date(t + 1)) > 0) scans.push({ id: Utilities.getUuid(), uid: p.uid, block: i % 2 ? 'Blok 6' : 'Blok 4', t: t, method: 'NFC', status: 'OK', team: p.team });
    }
  });
  TEAMS.forEach(tm => ingestScans_(scans.filter(s => s.team === tm), 'demo', tm));
}

/**
 * Peak-day load test: 50 pickers, about 7,500 scans spread over today's picking sessions,
 * sent the way the two phones send them (batches of 100). Times each batch and the dashboard.
 * Results go to the "Load test" tab. Run clearDemo afterwards to remove the test data.
 */
function loadTest() {
  setup();
  const today = day_(new Date());
  const ps = demoPickers_(50);
  const sess = sessions_();
  const scans = [];
  ps.forEach((p, i) => {
    const perPicker = 140 + (i * 13) % 41;                    // 140 to 180 bags each, about 7,500 in total
    let total = 0; sess.forEach(s => total += (at_(today, s.end) - at_(today, s.start)));
    const step = total / perPicker;
    let k = 0;
    sess.forEach(s => {
      for (let t = at_(today, s.start).getTime() + (i * 7919) % step; t < at_(today, s.end).getTime() && k < perPicker; t += step, k++) {
        scans.push({ id: Utilities.getUuid(), uid: p.uid, block: 'Blok ' + (1 + i % 8), t: Math.round(t), method: 'NFC', status: 'OK', team: p.team });
      }
    });
  });
  scans.sort((a, b) => a.t - b.t);
  const times = [];
  const t0 = Date.now();
  TEAMS.forEach(tm => {
    const mine = scans.filter(s => s.team === tm);
    for (let i = 0; i < mine.length; i += 100) {
      const b = Date.now();
      ingestScans_(mine.slice(i, i + 100), 'loadtest', tm);
      times.push(Date.now() - b);
    }
  });
  const ingestTotal = Date.now() - t0;
  const d0 = Date.now(); const d = dashboardData(today, 'hour'); const dashMs = Date.now() - d0;
  const s0 = Date.now(); ingestScans_(scans.slice(-100), 'loadtest-resend', TEAMS[0]); const resendMs = Date.now() - s0;
  const sh = sheet_('Load test', ['Run at', 'Scans', 'Batches', 'Avg batch ms', 'Slowest batch ms', 'All batches s', 'Dashboard ms', 'Resend 100 ms', 'Bags on dashboard', 'Taps rows']);
  const avg = Math.round(times.reduce((a, b) => a + b, 0) / times.length);
  sh.appendRow([new Date(), scans.length, times.length, avg, Math.max.apply(null, times), Math.round(ingestTotal / 1000), dashMs, resendMs, d.bags, taps_().getLastRow() - 1]);
  Logger.log('Scans %s, batches %s, avg %s ms, slowest %s ms, dashboard %s ms, bags %s', scans.length, times.length, avg, Math.max.apply(null, times), dashMs, d.bags);
}

function clearDemo() {
  [[pickers_(), 3], [taps_(), 6]].forEach(([sh, col]) => {
    const v = sh.getDataRange().getValues();
    const keep = v.filter((r, i) => i === 0 || String(r[col - 1]).indexOf(DEMO_PREFIX) !== 0);
    sh.clearContents();
    sh.getRange(1, 1, keep.length, keep[0].length).setValues(keep);
  });
  resetIdx_();
}
