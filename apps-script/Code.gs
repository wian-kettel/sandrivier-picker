/**
 * Sandrivier picker tracker: Google Sheet backend (v2, Oct 2026)
 *
 * The farm phone (scanning screen on GitHub Pages) sends scans here.
 * Everything is stored in this Google Sheet; nothing is stored on GitHub.
 *
 * Tabs (created or extended by setup()):
 *   Pickers  : Picker no | Name | Tag UID | Active | Registered | Type (Day/Contract)
 *   Taps     : Received | Date | Time | Picker no | Name | Tag UID | Block | Status |
 *              Team | Device | Method | Scan ID | Scan time
 *   Bins     : Date | Team | Plastic bins | Wooden bins | Updated
 *   Settings : Current block, Blocks, Farm PIN, Repeat window minutes,
 *              Kg per bag, Kg per plastic bin, Kg per wooden bin
 *
 * Status values: OK (counted), DUPLICATE (same band inside the repeat window, not counted),
 *                UNKNOWN (band not registered, not counted).
 * Method values: NFC (band tapped on the phone), MANUAL (number typed in), LINK (iPhone link test).
 */

const TZ = 'Africa/Johannesburg';
const SCAN_BACK_ROWS = 6000;
const DEMO_PREFIX = 'D0:00:00';
const TAP_HEADERS = ['Received', 'Date', 'Time', 'Picker no', 'Name', 'Tag UID', 'Block', 'Status',
                     'Team', 'Device', 'Method', 'Scan ID', 'Scan time'];
const PICKER_HEADERS = ['Picker no', 'Name', 'Tag UID', 'Active', 'Registered', 'Type'];
const BIN_HEADERS = ['Date', 'Team', 'Plastic bins', 'Wooden bins', 'Updated'];
const TEAMS = ['Team 1', 'Team 2'];

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
    ['Kg per bag', ''],
    ['Kg per plastic bin', ''],
    ['Kg per wooden bin', '']
  ];
  defaults.forEach(d => { if (have.indexOf(d[0]) < 0) s.appendRow(d); });
}

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.setFrozenRows(1); }
  const cur = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  if (cur.join('|') !== headers.join('|')) {
    // only fill in missing header cells, never rename existing columns
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

function checkPin_(pin) {
  const want = getSetting_('Farm PIN');
  if (!want || String(pin || '').trim() !== want) throw new Error('Wrong farm PIN.');
}

function pickerMap_() {
  const v = pickers_().getDataRange().getValues();
  const m = {};
  for (let i = 1; i < v.length; i++) {
    const uid = normUid_(v[i][2]);
    if (!uid) continue;
    m[uid] = {
      no: String(v[i][0]), name: String(v[i][1]), uid: uid,
      active: String(v[i][3]).toUpperCase() !== 'NO',
      type: String(v[i][5] || 'Day') === 'Contract' ? 'Contract' : 'Day'
    };
  }
  return m;
}

function readTaps_(maxRows) {
  const sh = taps_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const start = maxRows ? Math.max(2, last - maxRows + 1) : 2;
  return sh.getRange(start, 1, last - start + 1, TAP_HEADERS.length).getValues().map(r => {
    const rec = r[0] instanceof Date ? r[0] : new Date(r[0]);
    const st = r[12] instanceof Date ? r[12] : null;
    return {
      ts: st || rec, no: String(r[3]), name: String(r[4]), uid: normUid_(r[5]),
      block: String(r[6]), status: String(r[7]), team: String(r[8] || ''),
      device: String(r[9] || ''), method: String(r[10] || 'NFC'), id: String(r[11] || '')
    };
  }).filter(t => !isNaN(t.ts.getTime()));
}

/* ---------------- web entry points ---------------- */

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.api === 'roster') {
      checkPin_(p.pin);
      const m = pickerMap_();
      return json_({
        ok: true,
        pickers: Object.keys(m).map(k => m[k]),
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
    if (body.action === 'register') return json_({ ok: true, picker: registerPicker(body.uid, body.no, body.name, body.type) });
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
  lock.waitLock(20000);
  try {
    const m = pickerMap_();
    const byNo = {};
    Object.keys(m).forEach(k => { if (m[k].active) byNo[m[k].no] = m[k]; });
    const recent = readTaps_(SCAN_BACK_ROWS);
    const seen = {}, lastOk = {};
    recent.forEach(t => {
      if (t.id) seen[t.id] = true;
      if (t.status === 'OK' && (!lastOk[t.uid] || t.ts > lastOk[t.uid])) lastOk[t.uid] = t.ts;
    });
    const windowMs = repeatMs_();
    const now = new Date();
    const rows = [], accepted = [];
    scans.slice().sort((a, b) => a.t - b.t).forEach(s => {
      if (!s || !s.id) return;
      accepted.push(s.id);
      if (seen[s.id]) return;               // already stored (phone re-sent after a dropped reply)
      seen[s.id] = true;
      let uid = normUid_(s.uid);
      let p = uid ? m[uid] : null;
      if (!p && s.no && byNo[String(s.no)]) { p = byNo[String(s.no)]; uid = p.uid; }
      const st = new Date(Number(s.t) || now.getTime());
      let status;
      if (!p || !p.active) status = 'UNKNOWN';
      else if (s.status === 'DUPLICATE') status = 'DUPLICATE';
      else if (lastOk[uid] && st - lastOk[uid] < windowMs) status = 'DUPLICATE';
      else status = 'OK';
      if (status === 'OK') lastOk[uid] = st;
      rows.push([now, day_(st), Utilities.formatDate(st, TZ, 'HH:mm:ss'),
                 p ? p.no : String(s.no || ''), p ? p.name : '', uid, String(s.block || ''), status,
                 String(team || ''), String(device || ''), String(s.method || 'NFC'), String(s.id), st]);
    });
    if (rows.length) {
      const sh = taps_();
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, TAP_HEADERS.length).setValues(rows);
    }
    return { ok: true, accepted: accepted, stored: rows.length };
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- iPhone link scans (prototype, kept for testing) ---------------- */

function recordTap_(uid) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const p = pickerMap_()[uid];
    if (!p) return { state: 'unknown' };
    if (!p.active) return { state: 'inactive', picker: p };
    const now = new Date();
    const today = day_(now);
    let lastOk = null, count = 0;
    readTaps_(SCAN_BACK_ROWS).forEach(t => {
      if (t.uid !== uid || t.status !== 'OK') return;
      if (day_(t.ts) === today) count++;
      if (!lastOk || t.ts > lastOk) lastOk = t.ts;
    });
    const windowMs = repeatMs_();
    const ms = lastOk ? now - lastOk : Infinity;
    const status = ms < windowMs ? 'DUPLICATE' : 'OK';
    const block = getSetting_('Current block');
    taps_().appendRow([now, today, Utilities.formatDate(now, TZ, 'HH:mm:ss'), p.no, p.name, uid, block, status,
                       '', 'iPhone link', 'LINK', Utilities.getUuid(), now]);
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
    '.sub{font-size:18px;opacity:.9}input,select{font-size:20px;padding:12px;margin:6px 0;width:100%;max-width:320px;box-sizing:border-box;border-radius:8px;border:0}' +
    'button{font-size:20px;padding:14px 28px;margin-top:12px;border:0;border-radius:8px;background:#fff;color:#1d4ed8;font-weight:700}</style>';
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
  if (r.state === 'inactive') {
    return css + '<div class="p" style="background:#6b7280"><div class="name">Band disabled</div>' +
      '<div class="sub">' + esc_(r.picker.name) + '. Ask the office.</div></div>';
  }
  return css + '<div class="p" style="background:#1d4ed8" id="box">' +
    '<div class="sub">NEW BAND</div><div class="name">Link this band to a picker</div>' +
    '<div class="sub" style="margin-bottom:12px">Band ' + esc_(uid) + '</div>' +
    '<input id="no" placeholder="Picker number" inputmode="numeric">' +
    '<input id="nm" placeholder="Name and surname">' +
    '<select id="ty"><option>Day</option><option>Contract</option></select>' +
    '<button id="go" onclick="reg()">Register</button><div class="sub" id="msg" style="margin-top:12px"></div></div>' +
    '<script>function reg(){var no=document.getElementById("no").value,nm=document.getElementById("nm").value,ty=document.getElementById("ty").value;' +
    'document.getElementById("go").disabled=true;document.getElementById("msg").textContent="Saving...";' +
    'google.script.run.withSuccessHandler(function(p){var b=document.getElementById("box");b.style.background="#15803d";' +
    'b.innerHTML="<div class=name></div><div class=sub>Scan the band again to count a bag.</div>";b.firstChild.textContent="Linked: "+p.name+" (#"+p.no+")";})' +
    '.withFailureHandler(function(e){document.getElementById("go").disabled=false;document.getElementById("msg").textContent=e.message;})' +
    '.registerPicker(' + JSON.stringify(uid) + ',no,nm,ty);}</script>';
}

/* ---------------- called from pages or the phone ---------------- */

function registerPicker(uid, no, name, type) {
  uid = normUid_(uid); no = String(no || '').trim(); name = String(name || '').trim();
  type = type === 'Contract' ? 'Contract' : 'Day';
  if (!uid || !no || !name) throw new Error('Picker number and name are both required.');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const m = pickerMap_();
    if (m[uid]) throw new Error('This band is already linked to ' + m[uid].name + ' (#' + m[uid].no + ').');
    const clash = Object.keys(m).map(k => m[k]).filter(p => p.active && p.no === no);
    if (clash.length) throw new Error('Picker number ' + no + ' is already used by ' + clash[0].name + '.');
    const sh = pickers_();
    const row = sh.getLastRow() + 1;
    sh.getRange(row, 1).setNumberFormat('@');
    sh.getRange(row, 1, 1, PICKER_HEADERS.length).setValues([[no, name, uid, 'YES', new Date(), type]]);
    return { no: no, name: name, uid: uid, type: type, active: true };
  } finally {
    lock.releaseLock();
  }
}

function setBlock(block) {
  if (blocks_().indexOf(block) < 0) throw new Error('Unknown block: ' + block);
  setSetting_('Current block', block);
  return block;
}

function saveBins(date, team, plastic, wooden) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date || '')) throw new Error('Pick a day first.');
  if (TEAMS.indexOf(team) < 0) throw new Error('Pick a team.');
  const p = Math.max(0, Math.round(num_(plastic) || 0));
  const w = Math.max(0, Math.round(num_(wooden) || 0));
  const sh = bins_();
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (dstr_(v[i][0]) === date && String(v[i][1]) === team) {
      sh.getRange(i + 1, 3, 1, 3).setValues([[p, w, new Date()]]);
      return true;
    }
  }
  const row = sh.getLastRow() + 1;
  sh.getRange(row, 1).setNumberFormat('@');
  sh.getRange(row, 1, 1, 5).setValues([[date, team, p, w, new Date()]]);
  return true;
}

function dashboardData(dateStr, rankBy) {
  const date = /^\d{4}-\d{2}-\d{2}$/.test(dateStr || '') ? dateStr : day_(new Date());
  rankBy = rankBy === 'hour' ? 'hour' : 'bags';
  const all = readTaps_(0);
  const m = pickerMap_();
  const dayTaps = all.filter(t => day_(t.ts) === date);
  const ok = dayTaps.filter(t => t.status === 'OK');

  const per = {};
  ok.forEach(t => {
    const r = per[t.uid] || (per[t.uid] = { no: t.no, name: t.name, uid: t.uid, bags: 0, manual: 0, first: t.ts, last: t.ts, blocks: {}, team: t.team });
    r.bags++;
    if (t.method === 'MANUAL') r.manual++;
    if (t.ts < r.first) r.first = t.ts;
    if (t.ts > r.last) r.last = t.ts;
    r.blocks[t.block] = (r.blocks[t.block] || 0) + 1;
  });
  const rows = Object.keys(per).map(k => {
    const r = per[k];
    const hrs = (r.last - r.first) / 3600000;
    return {
      no: r.no, name: r.name, bags: r.bags, manual: r.manual,
      perHour: hrs >= 1 ? Math.round(r.bags / hrs * 10) / 10 : null,
      first: time_(r.first), last: time_(r.last), team: r.team || '',
      type: (m[r.uid] && m[r.uid].type) || 'Day',
      block: Object.keys(r.blocks).sort((a, b) => r.blocks[b] - r.blocks[a])[0]
    };
  });
  rows.sort((a, b) => b.bags - a.bags || a.name.localeCompare(b.name));

  let ranked;
  if (rankBy === 'hour') {
    ranked = rows.filter(r => r.perHour !== null).sort((a, b) => b.perHour - a.perHour || b.bags - a.bags);
  } else {
    ranked = rows.slice();
  }
  const pairs = [];
  for (let i = 0, j = ranked.length - 1; i < j; i++, j--) pairs.push({ strong: ranked[i], weak: ranked[j] });

  const byBlock = {}, hourly = {}, bagsByTeam = {};
  ok.forEach(t => {
    byBlock[t.block || '(none)'] = (byBlock[t.block || '(none)'] || 0) + 1;
    const h = Utilities.formatDate(t.ts, TZ, 'HH') + ':00';
    hourly[h] = (hourly[h] || 0) + 1;
    const tm = t.team || 'No team';
    bagsByTeam[tm] = (bagsByTeam[tm] || 0) + 1;
  });

  // 7-day leaderboard for rewards
  const days = {};
  for (let i = 0; i < 7; i++) days[day_(new Date(new Date(date + 'T12:00:00+02:00').getTime() - i * 86400000))] = true;
  const wk = {};
  all.forEach(t => {
    if (t.status !== 'OK' || !days[day_(t.ts)]) return;
    const r = wk[t.uid] || (wk[t.uid] = { no: t.no, name: t.name, bags: 0, days: {} });
    r.bags++; r.days[day_(t.ts)] = true;
  });
  const week = Object.keys(wk).map(k => ({
    no: wk[k].no, name: wk[k].name, bags: wk[k].bags, days: Object.keys(wk[k].days).length,
    type: (m[k] && m[k].type) || 'Day'
  })).sort((a, b) => b.bags - a.bags).slice(0, 10);

  // end-of-day check: estimated kg in bags vs estimated kg in bins
  const kgBag = num_(getSetting_('Kg per bag'));
  const kgP = num_(getSetting_('Kg per plastic bin'));
  const kgW = num_(getSetting_('Kg per wooden bin'));
  const binRows = bins_().getDataRange().getValues().slice(1).filter(r => dstr_(r[0]) === date);
  const teamNames = {};
  Object.keys(bagsByTeam).forEach(t => teamNames[t] = true);
  binRows.forEach(r => teamNames[String(r[1])] = true);
  const recon = Object.keys(teamNames).sort().map(tm => {
    const b = binRows.filter(r => String(r[1]) === tm);
    const plastic = b.reduce((s, r) => s + (Number(r[2]) || 0), 0);
    const wooden = b.reduce((s, r) => s + (Number(r[3]) || 0), 0);
    return reconRow_(tm, bagsByTeam[tm] || 0, plastic, wooden, kgBag, kgP, kgW, b.length > 0);
  });
  const tot = recon.reduce((a, r) => ({ bags: a.bags + r.bags, plastic: a.plastic + r.plastic, wooden: a.wooden + r.wooden, entered: a.entered || r.entered }),
                           { bags: 0, plastic: 0, wooden: 0, entered: false });
  const reconTotal = reconRow_('All teams', tot.bags, tot.plastic, tot.wooden, kgBag, kgP, kgW, tot.entered);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    date: date, rankBy: rankBy,
    total: ok.length,
    pickers: rows.length,
    registered: Object.keys(m).filter(k => m[k].active).length,
    duplicates: dayTaps.filter(t => t.status === 'DUPLICATE').length,
    unknown: dayTaps.filter(t => t.status === 'UNKNOWN').length,
    manual: ok.filter(t => t.method === 'MANUAL').length,
    avg: rows.length ? Math.round(ok.length / rows.length * 10) / 10 : 0,
    top: ranked.slice(0, 10),
    bottom: ranked.slice().reverse().slice(0, 10),
    rankedCount: ranked.length,
    rows: rows, pairs: pairs, byBlock: byBlock, hourly: hourly, week: week,
    recon: recon, reconTotal: reconTotal,
    weightsSet: !!(kgBag && kgP && kgW),
    weights: { bag: kgBag, plastic: kgP, wooden: kgW },
    teams: TEAMS,
    blocks: blocks_(),
    currentBlock: getSetting_('Current block'),
    sheetUrl: ss.getUrl(),
    xlsxUrl: 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx',
    updated: time_(new Date())
  };
}

function reconRow_(team, bags, plastic, wooden, kgBag, kgP, kgW, entered) {
  const kgIn = kgBag ? Math.round(bags * kgBag) : null;
  const kgOut = (kgP && kgW) ? Math.round(plastic * kgP + wooden * kgW) : null;
  const pct = (kgIn && kgOut !== null && entered) ? Math.round(kgOut / kgIn * 100) : null;
  return { team: team, bags: bags, plastic: plastic, wooden: wooden, kgIn: kgIn, kgOut: kgOut, pct: pct, entered: entered };
}

/* ---------------- demo data (run from the editor) ---------------- */

function seedDemo() {
  setup();
  const names = ['Anna', 'Piet', 'Sanna', 'Koos', 'Martha', 'Jan', 'Lena', 'Dawid', 'Rosa', 'Hendrik',
                 'Elsie', 'Sakkie', 'Grieta', 'Willem', 'Bettie', 'Frans', 'Mina', 'Klaas', 'Sophie', 'Gert',
                 'Maria', 'Jacob', 'Lettie', 'Abel'];
  const now = new Date();
  const today = day_(now);
  const start = new Date(today + 'T07:00:00+02:00');
  const end = now > start ? now : new Date(today + 'T13:00:00+02:00');
  const blocks = ['Blok 4', 'Blok 6'];
  const pRows = [], tRows = [];
  names.forEach((n, i) => {
    const uid = DEMO_PREFIX + ':' + ('0' + (i + 1).toString(16).toUpperCase()).slice(-2) + ':00:00:00';
    const no = String(100 + i + 1);
    const name = 'Demo ' + n;
    pRows.push([no, name, uid, 'YES', now, i % 3 === 0 ? 'Contract' : 'Day']);
    const minsPerBag = 6 + (i * 7) % 19;
    const block = blocks[i % 2];
    const team = TEAMS[i % 2];
    for (let t = start.getTime() + i * 60000; t < end.getTime(); t += minsPerBag * 60000) {
      const d = new Date(t);
      tRows.push([now, today, Utilities.formatDate(d, TZ, 'HH:mm:ss'), no, name, uid, block, 'OK',
                  team, 'demo', 'NFC', Utilities.getUuid(), d]);
    }
  });
  const ps = pickers_(), ts = taps_();
  ps.getRange(ps.getLastRow() + 1, 1, pRows.length, PICKER_HEADERS.length).setValues(pRows);
  if (tRows.length) ts.getRange(ts.getLastRow() + 1, 1, tRows.length, TAP_HEADERS.length).setValues(tRows);
}

function clearDemo() {
  [[pickers_(), 3], [taps_(), 6]].forEach(([sh, col]) => {
    const v = sh.getDataRange().getValues();
    const keep = v.filter((r, i) => i === 0 || String(r[col - 1]).indexOf(DEMO_PREFIX) !== 0);
    sh.clearContents();
    sh.getRange(1, 1, keep.length, keep[0].length).setValues(keep);
  });
}
