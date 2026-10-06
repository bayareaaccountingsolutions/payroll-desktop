/******************************************************************************
 *  DESKTOP APP API  —  add-on for the Payroll System (Code.gs v2.3)
 *  Add this as a SECOND file in the same Apps Script project (do not paste it
 *  into Code.gs). It lets the Payroll Desktop App (Windows / Mac) read and
 *  update THIS Google Sheet, so every staff member works on the same data.
 *
 *  ONE-TIME SETUP (about 5 minutes) — see README "Part A" for screenshots-level steps
 *   1. Extensions ▸ Apps Script ▸ "+" ▸ Script ▸ name it  DesktopApi  ▸ paste this file ▸ Save.
 *   2. In the function list choose  setupDesktopApp  ▸ Run ▸ allow permissions.
 *      A pop-up shows the ADMIN key and the STAFF key — keep them safe.
 *   3. Deploy ▸ New deployment ▸ type "Web app"
 *        Execute as:      Me
 *        Who has access:  Anyone
 *      ▸ Deploy ▸ copy the Web app URL (ends with /exec).
 *   4. In the desktop app ▸ Connect: paste the Web app URL + a key.
 *
 *  SECURITY: every request must carry the ADMIN or STAFF key. Without a key
 *  the URL returns nothing. Change keys any time: run  resetDesktopKeys.
 *   • STAFF key  : dashboard, attendance entry, payroll register, salary slips.
 *   • ADMIN key  : everything above + recalculate, sync, self-test,
 *                  finalize / unlock a month, email salary slips.
 *  Every change made from the app is written to the hidden "App Log" tab.
 *
 *  UPDATING THE SCRIPT LATER: replace Code.gs / DesktopApi.gs with the new
 *  versions shipped inside the app (Tools ▸ Sheet script), then
 *  Deploy ▸ Manage deployments ▸ ✏ Edit ▸ Version: New version ▸ Deploy.
 *  (Editing the existing deployment keeps the SAME URL — staff need not reconnect.)
 ******************************************************************************/

const DESKTOP_API_VERSION = '1.0.0';   // API contract the desktop app expects
const SHEET_SCRIPT_VERSION = '2.3';     // version of Code.gs this add-on was written for

/* ------------------------------ setup & keys ------------------------------ */
function setupDesktopApp() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('DESKTOP_ADMIN_KEY')) p.setProperty('DESKTOP_ADMIN_KEY', newKey_('ADM'));
  if (!p.getProperty('DESKTOP_STAFF_KEY')) p.setProperty('DESKTOP_STAFF_KEY', newKey_('STF'));
  showDesktopKeys();
}
function resetDesktopKeys() {
  const p = PropertiesService.getScriptProperties();
  p.setProperty('DESKTOP_ADMIN_KEY', newKey_('ADM'));
  p.setProperty('DESKTOP_STAFF_KEY', newKey_('STF'));
  showDesktopKeys('New keys created — the old keys stop working immediately. Give the new keys to your staff.');
}
function showDesktopKeys(note) {
  const p = PropertiesService.getScriptProperties();
  const msg = (note ? note + '\n\n' : '') +
    'ADMIN key (owner / HR head only):\n' + p.getProperty('DESKTOP_ADMIN_KEY') + '\n\n' +
    'STAFF key (attendance & slips):\n' + p.getProperty('DESKTOP_STAFF_KEY') + '\n\n' +
    'Next: Deploy ▸ New deployment ▸ Web app ▸ Execute as: Me ▸ Who has access: Anyone ▸ Deploy, then copy the /exec URL into the desktop app.';
  console.log(msg);   // always visible in Apps Script ▸ Execution log, even if no pop-up appears
}
/** Prints both keys in the Execution log at the bottom of the Apps Script editor (creates them if missing). */
function logDesktopKeys() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('DESKTOP_ADMIN_KEY')) p.setProperty('DESKTOP_ADMIN_KEY', newKey_('ADM'));
  if (!p.getProperty('DESKTOP_STAFF_KEY')) p.setProperty('DESKTOP_STAFF_KEY', newKey_('STF'));
  console.log('ADMIN key: ' + p.getProperty('DESKTOP_ADMIN_KEY'));
  console.log('STAFF key: ' + p.getProperty('DESKTOP_STAFF_KEY'));
}
function newKey_(prefix) { return prefix + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 24).toUpperCase(); }

/* ------------------------------- endpoints -------------------------------- */
function doGet() {
  return json_({ ok: true, app: 'Payroll System Desktop API', api: DESKTOP_API_VERSION, note: 'POST with an access key.' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'Bad request.' }); }
  const role = roleFor_(req.key);
  if (!role) return json_({ ok: false, error: 'Access key is not valid. Ask your administrator for the current key.', code: 'AUTH' });
  const action = String(req.action || ''), params = req.params || {}, user = String(req.user || 'Unknown').slice(0, 60);
  const def = API_ACTIONS_[action];
  if (!def) return json_({ ok: false, error: 'Unknown action: ' + action + '. Please update the sheet script (Tools ▸ Sheet script).', code: 'ACTION' });
  if (def.admin && role !== 'admin') return json_({ ok: false, error: 'This action needs the ADMIN key.', code: 'ROLE' });
  if (def.ready !== false) {
    const st = setupState_();
    if (st) return json_({ ok: false, error: st, code: 'SETUP' });
  }
  try {
    const data = def.fn(params, { role: role, user: user });
    if (def.log) appLog_(user, role, action, def.log(params, data));
    return json_({ ok: true, data: data });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function roleFor_(key) {
  if (!key) return null;
  const p = PropertiesService.getScriptProperties();
  if (key === p.getProperty('DESKTOP_ADMIN_KEY')) return 'admin';
  if (key === p.getProperty('DESKTOP_STAFF_KEY')) return 'staff';
  return null;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

const API_ACTIONS_ = {
  ping: { ready: false, fn: (p, c) => apiPing_(c) },
  summary: { fn: () => apiSummary_() },
  employees: { fn: () => apiEmployees_() },
  attendance: { fn: p => apiAttendance_(p) },
  saveAttendance: { fn: p => apiSaveAttendance_(p), log: (p, d) => `${p.month}: ${d.saved} cell(s) changed` },
  register: { fn: p => apiRegister_(p) },
  slipPdf: { fn: p => apiSlipPdf_(p), log: p => `Salary slip PDF ${p.id} ${p.month}` },
  registerPdf: { fn: p => apiRegisterPdf_(p), log: p => `Register PDF ${p.month}` },
  emailSlip: { admin: true, fn: p => apiEmailSlip_(p), log: (p, d) => `Emailed slip ${p.id} ${p.month} to ${d.to}` },
  recalc: { admin: true, fn: () => { recalcAll_(); renderRegister_(); return 'All 12 months recalculated.'; }, log: () => 'Recalculate all' },
  sync: { admin: true, fn: () => { syncEmployees(); markOffDays(); recalcAll_(); renderRegister_(); return 'Employees synced to month sheets.'; }, log: () => 'Sync employees' },
  markOffs: { admin: true, fn: () => { markOffDays(); recalcAll_(); return 'Weekly offs & holidays marked.'; }, log: () => 'Mark weekly offs' },
  selfTest: { admin: true, fn: () => runSelfTest(), log: () => 'Self-test' },
  finalize: { admin: true, fn: p => apiLock_(p, true), log: p => 'Finalize ' + p.month },
  unlock: { admin: true, fn: p => apiLock_(p, false), log: p => 'Unlock ' + p.month }
};

/* -------------------------------- actions --------------------------------- */
function apiPing_(c) {
  const ss = SpreadsheetApp.getActive();
  const st = ss.getSheetByName('Settings');
  return {
    role: c.role, api: DESKTOP_API_VERSION, scriptVersion: SHEET_SCRIPT_VERSION,
    setup: setupState_(), spreadsheetName: ss.getName(), spreadsheetUrl: ss.getUrl(),
    company: st ? String(st.getRange('C4').getValue()) : '', year: st ? Number(st.getRange('C8').getValue()) : null,
    capacity: APP.ROWS, timezone: ss.getSpreadsheetTimeZone()
  };
}

function apiBase_(ss) {
  const st = ss.getSheetByName('Settings');
  return {
    company: String(st.getRange('C4').getValue()), year: Number(st.getRange('C8').getValue()),
    currency: String(st.getRange('C9').getValue() || '₹'), currentMonth: String(st.getRange('C14').getValue()),
    statuses: st.getRange('L5:M9').getValues().map(r => ({ name: String(r[0]), pct: Number(r[1]) || 0 })),
    shifts: st.getRange('F5:F8').getValues().map(r => String(r[0])).filter(String)
  };
}

function apiSummary_() {
  const ss = SpreadsheetApp.getActive();
  const base = apiBase_(ss);
  const months = MONTHS.map(m => {
    const sh = ss.getSheetByName(m);
    const r = sh.getRange('A4:AF4').getValues()[0];
    const n = x => Number(x) || 0;
    return { month: m, locked: isLocked_(sh), active: n(r[0]), counts: [n(r[2]), n(r[3]), n(r[4]), n(r[5]), n(r[6])],
      attPct: r[7] === '' ? null : n(r[7]), paidPct: r[10] === '' ? null : n(r[10]), missing: n(r[13]),
      gross: n(r[16]), absence: n(r[19]), epf: n(r[22]), other: n(r[25]), net: n(r[28]) };
  });
  // today (live from the current month sheet: BP = expected today, BQ = today's entry)
  const tz = ss.getSpreadsheetTimeZone(), now = new Date();
  const ty = Number(Utilities.formatDate(now, tz, 'yyyy')), tm = Number(Utilities.formatDate(now, tz, 'M'));
  let today = null;
  if (ty === base.year) {
    const sh = ss.getSheetByName(MONTHS[tm - 1]);
    const v = sh.getRange(APP.M_FIRST, colN_('BP'), APP.ROWS, 2).getValues();
    const names = base.statuses.map(s => s.name), cnt = [0, 0, 0, 0, 0];
    let missing = 0, expected = 0;
    v.forEach(r => {
      const k = names.indexOf(String(r[1]).trim());
      if (k >= 0) cnt[k]++;
      if (Number(r[0]) === 1) { expected++; if (k < 0) missing++; }
    });
    today = { date: Utilities.formatDate(now, tz, 'dd-MMM-yyyy (EEE)'), counts: cnt, missing: missing, expected: expected };
  }
  const st = ss.getSheetByName('Settings');
  const m = st.getRange(APP.MASTER_FIRST, 1, APP.ROWS, 16).getValues().filter(r => String(r[0]).trim());
  const stat = r => String(r[14]).trim() || 'Active';
  const emp = { total: m.length, active: m.filter(r => stat(r) === 'Active').length,
    inactive: m.filter(r => stat(r) === 'Inactive').length, exited: m.filter(r => EXITED.indexOf(stat(r)) >= 0).length,
    byShift: {} };
  m.filter(r => stat(r) === 'Active').forEach(r => { const s = String(r[15] || r[4]); emp.byShift[s] = (emp.byShift[s] || 0) + 1; });
  return Object.assign(base, { months: months, today: today, employees: emp });
}

function apiEmployees_() {
  const ss = SpreadsheetApp.getActive();
  const st = ss.getSheetByName('Settings'), tz = ss.getSpreadsheetTimeZone();
  const heads = st.getRange('R23:X23').getValues()[0].map(String);
  const v = st.getRange(APP.MASTER_FIRST, 1, APP.ROWS, APP.EXIT_COL).getValues();
  const d = x => (x instanceof Date) ? Utilities.formatDate(x, tz, 'dd-MMM-yyyy') : '';
  const list = v.filter(r => String(r[0]).trim()).map(r => ({
    id: String(r[0]).trim(), name: String(r[1]), dept: String(r[2]), desig: String(r[3]), baseShift: String(r[4]),
    join: d(r[5]), gross: Number(r[6]) || 0, method: String(r[7]), epf: String(r[9]), email: String(r[13]),
    status: String(r[14]).trim() || 'Active', shift: String(r[15]), salary: Number(r[16]) || 0,
    heads: heads.map((h, i) => ({ name: h, amount: Number(r[17 + i]) || 0 })), esi: Number(r[24]) || 0,
    inHand: Number(r[25]) || 0, exit: d(r[26])
  }));
  return { heads: heads, employees: list, currency: String(st.getRange('C9').getValue() || '₹') };
}

function monthSheet_(name) {
  const m = monthFromText_(name);
  if (!m) throw new Error('Not a month name: ' + name);
  return SpreadsheetApp.getActive().getSheetByName(m);
}

function apiAttendance_(p) {
  const ss = SpreadsheetApp.getActive(), sh = monthSheet_(p.month), tz = ss.getSpreadsheetTimeZone();
  const base = apiBase_(ss);
  const F = APP.M_FIRST, N = APP.ROWS;
  const dim = Number(sh.getRange('F2').getValue());
  const days = sh.getRange('H7:AL8').getValues();
  const v = sh.getRange(F, 1, N, colN_('CZ')).getValues();
  const exitV = sh.getMaxColumns() >= XC.EXIT ? sh.getRange(F, XC.EXIT, N, 1).getValues() : v.map(() => ['']);
  const hidden = []; for (let i = 0; i < N; i++) hidden.push(sh.isRowHiddenByUser(F + i));
  const ser = x => toSerial_(x, tz);
  const rows = [];
  v.forEach((r, i) => {
    const id = String(r[0]).trim(); if (!id || hidden[i]) return;
    rows.push({ row: F + i, id: id, name: String(r[1]), shift: String(r[4]),
      att: r.slice(7, 38).map(x => String(x).trim()), types: r.slice(colN_('BV') - 1, colN_('CZ')).map(String),
      join: ser(r[colN_('BG') - 1]), exit: ser(exitV[i][0]), other: r[colN_('AY') - 1],
      missing: Number(r[colN_('AR') - 1]) || 0, net: Number(r[colN_('BE') - 1]) || 0, remarks: String(r[colN_('BF') - 1]) });
  });
  return { month: sh.getName(), year: base.year, locked: isLocked_(sh), dim: dim, statuses: base.statuses, currency: base.currency,
    today: ser(new Date()), trackStart: ser(ss.getSheetByName('Settings').getRange('C13').getValue()),
    dates: days[0].map(x => ser(x)), dayNames: days[1].map(x => String(x).slice(0, 3)), rows: rows };
}

/** params: {month, changes:[{row, day, value}], other:[{row, amount}]} */
function apiSaveAttendance_(p) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActive(), sh = monthSheet_(p.month);
    if (isLocked_(sh)) throw new Error(sh.getName() + ' is finalized (locked). Ask an admin to unlock it first.');
    const mi = MONTHS.indexOf(sh.getName());
    const names = named_(ss, 'StatusNames').getValues().map(r => String(r[0]));
    const dim = Number(sh.getRange('F2').getValue());
    const okRow = r => Number.isInteger(r) && r >= APP.M_FIRST && r <= APP.M_LAST;
    const touched = {}; let saved = 0, minDay = 32, maxDay = 0;
    (p.changes || []).forEach(c => {
      const r = Number(c.row), d = Number(c.day), val = String(c.value || '').trim();
      if (!okRow(r) || !(d >= 1 && d <= dim)) throw new Error('Invalid cell (row ' + c.row + ', day ' + c.day + ').');
      if (val && names.indexOf(val) < 0) throw new Error('Not an attendance status: ' + val);
      sh.getRange(r, 7 + d).setValue(val); touched[r] = true; saved++;
      minDay = Math.min(minDay, d); maxDay = Math.max(maxDay, d);
    });
    (p.other || []).forEach(o => {
      const r = Number(o.row), a = (o.amount === '' || o.amount === null) ? '' : Number(o.amount);
      if (!okRow(r) || (a !== '' && !(a >= 0))) throw new Error('Invalid Other Earnings amount.');
      sh.getRange(r, colN_('AY')).setValue(a); touched[r] = true; saved++;
    });
    const rows = Object.keys(touched).map(Number);
    if (rows.length) {
      const ids = rows.map(r => String(sh.getRange(r, 1).getValue()).trim()).filter(String);
      const months = [mi + 1];                                   // same neighbour rule as the sheet's edit trigger
      if (minDay <= 10 && mi > 0) months.unshift(mi);
      if (maxDay >= 22 && mi < 11) months.push(mi + 2);
      if (ids.length) recalcMonths_(months, ids);
    }
    return { saved: saved, attendance: apiAttendance_({ month: p.month }) };
  } finally { lock.releaseLock(); }
}

function apiRegister_(p) {
  const ss = SpreadsheetApp.getActive(), sh = monthSheet_(p.month), base = apiBase_(ss);
  const F = APP.M_FIRST, N = APP.ROWS;
  const v = sh.getRange(F, 1, N, Math.min(sh.getMaxColumns(), OUT_LAST)).getValues();
  const g = (r, c) => r[colN_(c) - 1];
  const n = x => Number(x) || 0;
  const rows = v.filter(r => String(r[0]).trim() && (String(g(r, 'BI')) === 'Active' || n(g(r, 'BL')) > 0)).map(r => ({
    id: String(r[0]).trim(), name: String(r[1]), dept: String(r[2]), shift: String(r[4]), rate: n(r[5]),
    workDays: n(g(r, 'BO')), counts: ['AM', 'AN', 'AO', 'AP', 'AQ'].map(c => n(g(r, c))), sandwich: n(g(r, 'EG')),
    missing: n(g(r, 'AR')), paidDays: n(g(r, 'AT')), unpaid: n(g(r, 'AU')), earned: n(g(r, 'AX')), otherEarn: n(g(r, 'AY')),
    gross: n(g(r, 'AZ')), lop: n(g(r, 'BA')), epfEe: n(g(r, 'BR')), epfEr: n(g(r, 'BS')), esi: n(g(r, 'EF')),
    otherDed: n(g(r, 'BC')), advance: n(r[XC.ADV - 1]), totalDed: n(g(r, 'BD')), net: n(g(r, 'BE')), remarks: String(g(r, 'BF'))
  }));
  return { month: sh.getName(), year: base.year, locked: isLocked_(sh), statuses: base.statuses.map(s => s.name),
    currency: base.currency, company: base.company, rows: rows };
}

function slipLabel_(ss, id) {
  const list = named_(ss, 'ActiveList').getValues().map(r => String(r[0])).filter(String);
  const hit = list.find(x => x.split(' - ')[0].trim() === id);
  if (hit) return hit;
  const ids = named_(ss, 'Emp_ID').getValues().map(r => String(r[0]).trim());
  const k = ids.indexOf(id); if (k < 0) throw new Error('Employee ' + id + ' is not in the Employee Master.');
  return id + ' - ' + named_(ss, 'Emp_Name').getValues()[k][0];
}

/** Fills the Salary Slip tab for (id, month), runs fn(info), then restores the previous selection. */
function withSlip_(p, fn) {
  const lock = LockService.getScriptLock(); lock.waitLock(60000);
  const ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName('Salary Slip');
  const c0 = sh.getRange('C2').getValue(), f0 = sh.getRange('F2').getValue();
  try {
    const month = monthFromText_(p.month); if (!month) throw new Error('Choose a month.');
    sh.getRange('C2').setValue(slipLabel_(ss, String(p.id).trim()));
    sh.getRange('F2').setValue(month);
    SpreadsheetApp.flush();
    const msg = String(sh.getRange('K6').getValue());
    if (msg !== 'OK') throw new Error(msg);
    const st = ss.getSheetByName('Settings'), idx = Number(sh.getRange('K8').getValue());
    const info = { ss: ss, sheet: sh, id: String(sh.getRange('K2').getValue()), name: String(sh.getRange('F13').getDisplayValue()),
      month: month, year: String(sh.getRange('F10').getDisplayValue()), net: String(sh.getRange('F32').getDisplayValue()),
      email: idx ? String(st.getRange(APP.MASTER_FIRST + idx - 1, 14).getValue()).trim() : '', company: String(st.getRange('C4').getValue()) };
    return fn(info);
  } finally {
    sh.getRange('C2').setValue(c0); sh.getRange('F2').setValue(f0); SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

function apiSlipPdf_(p) {
  return withSlip_(p, info => {
    const blob = exportSlipPdf_(info);
    if (p.saveToDrive) slipFolder_().createFile(blob.copyBlob());
    return { fileName: blob.getName(), base64: Utilities.base64Encode(blob.getBytes()), net: info.net, email: info.email };
  });
}

function apiEmailSlip_(p) {
  return withSlip_(p, info => {
    if (!info.email) throw new Error('No email for ' + info.name + ' in Settings ▸ Employee Master (column N).');
    const blob = exportSlipPdf_(info);
    slipFolder_().createFile(blob.copyBlob());
    MailApp.sendEmail({ to: info.email, subject: `Salary Slip — ${info.month} ${info.year} — ${info.company}`,
      body: `Dear ${info.name},\n\nPlease find attached your salary slip for ${info.month} ${info.year}.\n\nRegards,\n${info.company}`,
      name: info.company, attachments: [blob] });
    return { to: info.email, name: info.name };
  });
}

function apiRegisterPdf_(p) {
  const lock = LockService.getScriptLock(); lock.waitLock(60000);
  const ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(REG);
  const c0 = sh.getRange('C2').getValue();
  try {
    const month = monthFromText_(p.month); if (!month) throw new Error('Choose a month.');
    sh.getRange('C2').setValue(month); renderRegister_(); SpreadsheetApp.flush();
    const n = Number(sh.getRange('A5').getValue()) || 0;
    if (!n) throw new Error('No employees / attendance for ' + month + '.');
    const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${sh.getSheetId()}` +
      '&size=A4&portrait=false&scale=2&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=CENTER' +
      '&fzr=true&fzc=false&top_margin=0.4&bottom_margin=0.4&left_margin=0.3&right_margin=0.3' + `&r1=0&c1=0&r2=${8 + n}&c2=27`;
    const resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) throw new Error('PDF export failed (HTTP ' + resp.getResponseCode() + '). Please try again.');
    const name = `Payroll Register - ${month} ${sh.getRange('F2').getDisplayValue()}.pdf`;
    return { fileName: name, base64: Utilities.base64Encode(resp.getBlob().getBytes()) };
  } finally {
    sh.getRange('C2').setValue(c0); renderRegister_(); lock.releaseLock();
  }
}

function apiLock_(p, lockIt) {
  const sh = monthSheet_(p.month);
  if (lockIt) { if (isLocked_(sh)) return sh.getName() + ' is already finalized.'; lockSheet_(sh); return '🔒 ' + sh.getName() + ' finalized.'; }
  if (!isLocked_(sh)) return sh.getName() + ' is not locked.';
  unlockSheet_(sh); return '🔓 ' + sh.getName() + ' unlocked and recalculated.';
}

/* --------------------------------- log ------------------------------------ */
function appLog_(user, role, action, detail) {
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName('App Log');
    if (!sh) {
      sh = ss.insertSheet('App Log');
      sh.getRange('A1:E1').setValues([['Time', 'User', 'Role', 'Action', 'Detail']]).setFontWeight('bold');
      sh.setColumnWidths(1, 5, 160); sh.hideSheet();
    }
    sh.appendRow([new Date(), user, role, action, String(detail || '')]);
  } catch (e) { /* logging never blocks the request */ }
}
