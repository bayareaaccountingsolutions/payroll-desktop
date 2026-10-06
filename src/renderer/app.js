/* Payroll System — desktop app UI (renderer). All data comes live from the shared Google Sheet. */
'use strict';
const P = window.payroll;
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const $ = s => document.querySelector(s);
const h = s => String(s === undefined || s === null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const S = { cfg: {}, ping: null, role: null, view: 'dashboard', cache: {}, month: null, att: null, pend: {}, sel: null, cur: '₹' };

/* ------------------------------- helpers -------------------------------- */
const inr = new Intl.NumberFormat('en-IN', { maximumFractionDigits: 0 });
const money = n => (n < 0 ? '-' : '') + S.cur + inr.format(Math.abs(Math.round(Number(n) || 0)));
const num = n => { const x = Math.round((Number(n) || 0) * 100) / 100; return String(x); };
const pct = n => (n === null || n === undefined || n === '') ? '—' : (Number(n) * 100).toFixed(1) + '%';
function toast(msg, err) {
  const d = document.createElement('div'); if (err) d.className = 'err'; d.textContent = msg;
  $('#toast').appendChild(d); setTimeout(() => d.remove(), err ? 9000 : 4500);
}
async function call(action, params) {
  const r = await P.api(action, params);
  if (!r || !r.ok) {
    const msg = (r && r.error) || 'Something went wrong.';
    if (r && (r.code === 'NET' || r.code === 'AUTH' || r.code === 'DEPLOY')) setConn(false, r.code === 'NET' ? 'Offline' : 'Not authorised');
    throw new Error(msg);
  }
  setConn(true);
  return r.data;
}
function setConn(ok, label) {
  const c = $('#conn'); c.className = 'pill ' + (ok ? 'ok' : 'bad'); c.innerHTML = `<span class="dot"></span>${ok ? 'Connected' : h(label || 'Offline')}`;
}
function loading(msg) { $('#view').innerHTML = `<div class="empty"><span class="spin"></span> ${h(msg || 'Loading from Google Sheet…')}</div>`; }
function failed(err, retry) {
  $('#view').innerHTML = `<div class="empty"><b>Could not load</b>${h(err.message || err)}<br><br>
    <button class="btn pri" id="retry">Try again</button> <button class="btn" id="goset">Settings</button></div>`;
  $('#retry').onclick = retry; $('#goset').onclick = () => go('settings');
}
const monthOpts = sel => MONTHS.map(m => `<option ${m === sel ? 'selected' : ''}>${m}</option>`).join('');
const defMonth = () => S.month || (S.cache.summary && S.cache.summary.currentMonth) || MONTHS[new Date().getMonth()];
const isAdmin = () => S.role === 'admin';
const codeOf = name => { const w = String(name).trim().split(/\s+/); return w.length > 1 ? w.map(x => x[0]).join('').toUpperCase() : (name.length <= 3 ? name.toUpperCase() : name[0].toUpperCase()); };

/* ------------------------------- routing -------------------------------- */
const VIEWS = {
  dashboard: ['Dashboard', viewDashboard], attendance: ['Attendance', viewAttendance], register: ['Payroll Register', viewRegister],
  slip: ['Salary Slip', viewSlip], employees: ['Employees', viewEmployees], sheet: ['Google Sheet', viewSheet],
  tools: ['Admin Tools', viewTools], settings: ['Settings', viewSettings]
};
function go(v) {
  if (S.view === 'attendance' && v !== 'attendance' && Object.keys(S.pend).length &&
      !confirm('You have unsaved attendance changes. Leave without saving?')) return;
  if (v !== 'attendance') S.pend = {};
  hidePop();
  S.view = v;
  document.querySelectorAll('#nav button').forEach(b => b.classList.toggle('on', b.dataset.v === v));
  $('#title').innerHTML = h(VIEWS[v][0]) + (S.ping ? `<span class="sub">${h(S.ping.company)} · ${h(S.ping.year)}</span>` : '');
  if (!S.ping && v !== 'settings') { viewConnectFirst(); return; }
  VIEWS[v][1]();
}
document.querySelectorAll('#nav button').forEach(b => b.onclick = () => go(b.dataset.v));

/* ------------------------------- startup -------------------------------- */
async function boot() {
  S.cfg = await P.getConfig();
  if (S.cfg.theme) document.documentElement.dataset.theme = S.cfg.theme;
  $('#ver').textContent = 'v' + S.cfg.appVersion;
  if (S.cfg.demo) banner('warn', 'DEMO MODE — sample data, not connected to Google.');
  else if (!S.cfg.webAppUrl || !S.cfg.key) { go('settings'); return; }
  loading('Connecting to your Google Sheet…');
  try { await connect(); go('dashboard'); }
  catch (e) { failed(e, boot); }
}
async function connect() {
  const p = await call('ping');
  S.ping = p; S.role = p.role;
  $('#brandName').textContent = p.company || 'Payroll System';
  $('#who').innerHTML = `${h(S.cfg.userName || 'User')} <span class="pill ${p.role === 'admin' ? 'admin' : ''}" style="margin-left:4px">${p.role === 'admin' ? 'Admin' : 'Staff'}</span>`;
  document.querySelectorAll('#nav [data-admin]').forEach(b => b.style.display = p.role === 'admin' ? '' : 'none');
  if (p.spreadsheetUrl && p.spreadsheetUrl !== S.cfg.sheetUrl) S.cfg = await P.setConfig({ sheetUrl: p.spreadsheetUrl });
  if (p.setup) banner('warn', '⚠ ' + p.setup);
  const sf = await P.scriptFiles();
  if (sf.scriptVersion && p.scriptVersion && sf.scriptVersion !== p.scriptVersion && isAdmin())
    banner('info', `The sheet script is v${p.scriptVersion}; this app includes v${sf.scriptVersion}. Admin Tools ▸ Sheet script shows how to update it (2 minutes, data is kept).`);
}
function banner(kind, html, actions) {
  const b = $('#banner'); b.className = 'show ' + kind; b.innerHTML = `<span style="flex:1">${html}</span>` + (actions || '') + '<button class="btn ghost sm" id="bx">✕</button>';
  $('#bx').onclick = () => { b.className = ''; };
}
function viewConnectFirst() {
  $('#view').innerHTML = `<div class="empty"><b>Not connected to your Google Sheet</b>Open Settings and paste the Web App URL and your access key.<br><br>
    <button class="btn pri" id="gs">Open Settings</button></div>`;
  $('#gs').onclick = () => go('settings');
}

/* ------------------------------- updates -------------------------------- */
P.onUpdate(s => {
  S.upd = s;
  if (s.state === 'ready') banner('ok', `✔ Version ${h(s.version)} is ready. Restart to finish updating (your data is in Google Sheets and is not affected).`,
    '<button class="btn pri sm" id="ui">Restart & update</button>');
  else if (s.state === 'available-manual') banner('info', `A new version ${h(s.version)} is available.`, '<button class="btn pri sm" id="ud">Download</button>');
  else if (s.state === 'downloading' && s.percent === undefined) toast(`Downloading update ${s.version || ''}…`);
  const i = $('#ui'); if (i) i.onclick = () => P.installUpdate();
  const d = $('#ud'); if (d) d.onclick = () => P.openLink(s.url);
  if (S.view === 'settings') { const u = $('#updState'); if (u) u.innerHTML = updText(s); }
});
function updText(s) {
  if (!s) return 'Not checked yet.';
  return ({ dev: 'Updates work in the installed app (not in developer mode).', checking: '<span class="spin"></span> Checking…', none: '✔ You have the latest version.',
    downloading: `<span class="spin"></span> Downloading ${s.percent !== undefined ? s.percent + '%' : ''}…`, ready: `✔ Version ${h(s.version)} downloaded — restart to install.`,
    'available-manual': `Version ${h(s.version)} is available — click Download.`, error: 'Could not check: ' + h(s.message) })[s.state] || '';
}

/* ------------------------------- dashboard ------------------------------ */
async function viewDashboard(force) {
  if (!S.cache.summary || force) {
    loading();
    try { S.cache.summary = await call('summary'); } catch (e) { return failed(e, () => viewDashboard(true)); }
  }
  const d = S.cache.summary; S.cur = d.currency || '₹';
  const sel = defMonth(), mrow = d.months.find(m => m.month === sel) || d.months[0];
  const st = d.statuses, col = i => `var(--t${i})`;
  const t = d.today;
  const kp = (l, v, c, cls) => `<div class="kpi"><div class="l">${h(l)}</div><div class="v ${cls || ''}" ${c ? `style="color:${c}"` : ''}>${v}</div></div>`;
  $('#view').innerHTML = `
    <div class="toolbar"><div class="sec" style="margin:0">Today ${t ? '· ' + h(t.date) : ''}</div><div class="grow"></div>
      <button class="btn" id="rf">↻ Refresh</button></div>
    <div class="grid g4">
      ${kp('Active employees', d.employees.active)}
      ${t ? st.map((s, i) => kp(s.name + ' today', t.counts[i], col(i))).join('') : kp('Today', '—')}
      ${kp('Not marked yet', t ? t.missing : '—', '', 'red')}
      ${kp('Marked today', t ? `${t.expected - t.missing} / ${t.expected}` : '—')}
    </div>
    <div class="toolbar" style="margin-top:22px"><div class="sec" style="margin:0">Payroll</div>
      <select id="dm">${monthOpts(sel)}</select>${mrow.locked ? '<span class="pill lock">🔒 Finalized</span>' : ''}</div>
    <div class="grid g4">
      ${kp('Gross payroll', money(mrow.gross))}${kp('Absence deduction', money(mrow.absence), col(4))}
      ${kp('EPF (total)', money(mrow.epf))}${kp('Other deductions', money(mrow.other))}
      ${kp('Net payroll', money(mrow.net), '', 'red')}${kp('Attendance %', pct(mrow.attPct))}
      ${kp('Paid attendance %', pct(mrow.paidPct))}${kp('Missing entries', mrow.missing, '', mrow.missing ? 'red' : '')}
    </div>
    <div class="grid g2" style="margin-top:22px">
      <div class="card chart"><h3>Net payroll by month · ${h(d.year)}</h3>${barChart(d.months.map(m => ({ k: m.month, v: m.net })), sel)}</div>
      <div class="card chart"><h3>Attendance % by month · ${h(d.year)}</h3>${lineChart(d.months.map(m => ({ k: m.month, v: m.attPct })))}</div>
    </div>
    <div class="sec">Monthly overview</div>
    <div class="scroll" style="max-height:none"><table class="t"><thead><tr><th>Month</th><th class="n">Active</th>
      ${st.map(s => `<th class="n">${h(s.name)}</th>`).join('')}<th class="n">Missing</th><th class="n">Att. %</th><th class="n">Paid %</th>
      <th class="n">Gross</th><th class="n">Absence</th><th class="n">EPF</th><th class="n">Other</th><th class="n">Net</th><th></th></tr></thead><tbody>
      ${d.months.map(m => `<tr ${m.month === d.currentMonth ? 'style="font-weight:700"' : ''}><td>${h(m.month)}</td><td class="n">${m.active}</td>
        ${m.counts.map(c => `<td class="n">${num(c)}</td>`).join('')}<td class="n" ${m.missing ? 'style="color:var(--t4);font-weight:700"' : ''}>${m.missing}</td>
        <td class="n">${pct(m.attPct)}</td><td class="n">${pct(m.paidPct)}</td><td class="n">${money(m.gross)}</td><td class="n">${money(m.absence)}</td>
        <td class="n">${money(m.epf)}</td><td class="n">${money(m.other)}</td><td class="n net">${money(m.net)}</td><td>${m.locked ? '🔒' : ''}</td></tr>`).join('')}
      </tbody><tfoot><tr><td>Total</td><td></td>${st.map((s, i) => `<td class="n">${num(d.months.reduce((a, m) => a + m.counts[i], 0))}</td>`).join('')}
        <td class="n">${d.months.reduce((a, m) => a + m.missing, 0)}</td><td></td><td></td>
        ${['gross', 'absence', 'epf', 'other', 'net'].map(k => `<td class="n">${money(d.months.reduce((a, m) => a + m[k], 0))}</td>`).join('')}<td></td></tr></tfoot></table></div>`;
  $('#rf').onclick = () => viewDashboard(true);
  $('#dm').onchange = e => { S.month = e.target.value; viewDashboard(); };
  bindTips();
}
function barChart(data, sel) {
  const W = 560, H = 230, L = 64, B = 26, T = 10, max = niceMax(Math.max(1, ...data.map(x => x.v)));
  const bw = (W - L) / data.length, y = v => T + (H - T - B) * (1 - v / max);
  let s = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Net payroll by month">`;
  for (let i = 0; i <= 4; i++) { const v = max * i / 4; s += `<line class="gl" x1="${L}" x2="${W}" y1="${y(v)}" y2="${y(v)}"/><text class="ax" x="${L - 6}" y="${y(v) + 3}" text-anchor="end">${short(v)}</text>`; }
  data.forEach((d, i) => {
    const x = L + i * bw + bw * 0.2, w = bw * 0.6, top = y(Math.max(0, d.v)), hgt = Math.max(0, y(0) - top);
    if (hgt > 0) s += `<path class="bar ${d.k === sel ? 'cur' : ''}" d="M${x},${y(0)} V${top + 4} q0,-4 4,-4 h${w - 8} q4,0 4,4 V${y(0)} Z"/>`;
    s += `<rect class="hit" x="${L + i * bw}" y="${T}" width="${bw}" height="${H - T - B}" data-tip="${h(d.k)}: ${h(money(d.v))}"/>`;
    s += `<text class="ax" x="${x + w / 2}" y="${H - 8}" text-anchor="middle">${d.k.slice(0, 3)}</text>`;
  });
  return s + '</svg>';
}
function lineChart(data) {
  const W = 560, H = 230, L = 44, B = 26, T = 10, bw = (W - L) / data.length, y = v => T + (H - T - B) * (1 - v);
  let s = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Attendance percentage by month">`;
  for (let i = 0; i <= 4; i++) s += `<line class="gl" x1="${L}" x2="${W}" y1="${y(i / 4)}" y2="${y(i / 4)}"/><text class="ax" x="${L - 6}" y="${y(i / 4) + 3}" text-anchor="end">${i * 25}%</text>`;
  const pts = data.map((d, i) => d.v === null ? null : [L + i * bw + bw / 2, y(Math.min(1, d.v))]);
  let path = '', on = false;
  pts.forEach(p => { if (!p) { on = false; return; } path += (on ? 'L' : 'M') + p[0].toFixed(1) + ',' + p[1].toFixed(1); on = true; });
  s += `<path class="ln" d="${path}"/>`;
  data.forEach((d, i) => {
    if (pts[i]) s += `<circle class="pt" cx="${pts[i][0]}" cy="${pts[i][1]}" r="4"/>`;
    s += `<rect class="hit" x="${L + i * bw}" y="${T}" width="${bw}" height="${H - T - B}" data-tip="${h(d.k)}: ${pct(d.v)}"/>`;
    s += `<text class="ax" x="${L + i * bw + bw / 2}" y="${H - 8}" text-anchor="middle">${d.k.slice(0, 3)}</text>`;
  });
  return s + '</svg>';
}
function niceMax(v) { const p = Math.pow(10, Math.floor(Math.log10(v))); return Math.ceil(v / p * 2) / 2 * p; }
function short(v) { return v >= 1e7 ? (v / 1e7).toFixed(1) + 'Cr' : v >= 1e5 ? (v / 1e5).toFixed(1) + 'L' : v >= 1e3 ? Math.round(v / 1e3) + 'K' : String(Math.round(v)); }
function bindTips() {
  const tip = $('#tip');
  document.querySelectorAll('[data-tip]').forEach(el => {
    el.onmousemove = e => { tip.textContent = el.dataset.tip; tip.style.display = 'block'; tip.style.left = (e.clientX + 12) + 'px'; tip.style.top = (e.clientY - 30) + 'px'; };
    el.onmouseleave = () => { tip.style.display = 'none'; };
  });
}

/* ------------------------------- attendance ----------------------------- */
async function viewAttendance(force) {
  const m = defMonth();
  if (!S.att || S.att.month !== m || force) {
    loading('Loading ' + m + ' attendance…');
    try { S.att = await call('attendance', { month: m }); S.pend = {}; } catch (e) { return failed(e, () => viewAttendance(true)); }
  }
  renderAttendance();
}
function cellInfo(a, r, d) {
  const s = a.dates[d - 1];
  if (!s || d > a.dim) return { cls: 'd ro na', txt: '', ro: true };
  if (s < r.join || (r.exit && s > r.exit)) return { cls: 'd ro pre', txt: '', ro: true };
  const key = r.row + '|' + d, val = key in S.pend ? S.pend[key] : r.att[d - 1];
  const k = a.statuses.findIndex(x => x.name === val), type = r.types[d - 1];
  const pend = key in S.pend ? ' pend' : '';
  if (k >= 0) return { cls: `d st${k}${pend}${a.locked ? ' ro' : ''}`, txt: codeOf(val), ro: a.locked, tip: val };
  if (type === 'Off') return { cls: 'd ro off', txt: 'Off', ro: true, tip: 'Weekly off' };
  if (type === 'Hol') return { cls: 'd ro hol', txt: 'Hol', ro: true, tip: 'Holiday' };
  const miss = s < a.today && s >= a.trackStart;
  return { cls: `d${miss ? ' miss' : ''}${pend}${a.locked ? ' ro' : ''}`, txt: '', ro: a.locked, tip: miss ? 'MISSING — not marked' : 'Not marked' };
}
function renderAttendance() {
  const a = S.att, q = (S.attQ || '').toLowerCase();
  const rows = a.rows.filter(r => !q || (r.id + ' ' + r.name + ' ' + r.shift).toLowerCase().includes(q));
  const np = Object.keys(S.pend).length;
  const todayIdx = a.dates.indexOf(a.today);
  $('#view').innerHTML = `
    <div class="toolbar"><select id="am">${monthOpts(a.month)}</select>
      <input type="search" id="aq" placeholder="Search employee…" value="${h(S.attQ || '')}" style="width:200px">
      ${a.locked ? '<span class="pill lock">🔒 Finalized — read only</span>' : ''}
      ${!a.locked && todayIdx >= 0 ? '<button class="btn" id="mt">Mark blanks today as ' + h(a.statuses[0].name) + '</button>' : ''}
      <div class="grow"></div>
      <button class="btn" id="ar">↻ Refresh</button>
      <button class="btn ghost" id="ad" ${np ? '' : 'disabled'}>Discard</button>
      <button class="btn pri" id="as" ${np ? '' : 'disabled'}>Save ${np ? '(' + np + ')' : ''}</button></div>
    <div class="legend">${a.statuses.map((s, i) => `<span class="st${i}" style="background:var(--s${i});color:var(--t${i})">${h(codeOf(s.name))} = ${h(s.name)} · ${Math.round(s.pct * 100)}% paid · key ${i + 1}${'/' + codeOf(s.name)[0]}</span>`).join('')}
      <span style="background:var(--miss)">MISSING</span><span style="background:var(--off)">Weekly off</span><span style="background:var(--hol)">Holiday</span>
      <span style="background:var(--pre)">Not employed</span><span class="hint" style="padding:3px 4px">Click a cell or use keys · Delete clears · arrows move</span></div>
    <div class="att-wrap"><table class="att"><thead><tr><th class="sticky">ID</th><th class="sticky2">Employee</th>
      ${a.dates.map((s, i) => i < a.dim ? `<th class="${s === a.today ? 'today' : ''}">${i + 1}<small>${h(a.dayNames[i])}</small></th>` : '').join('')}
      <th>Missing</th><th>Net</th></tr></thead><tbody>
      ${rows.length ? rows.map(r => `<tr><td class="id sticky">${h(r.id)}</td><td class="nm sticky2" title="${h(r.remarks)}">${h(r.name)}<small>${h(r.shift)}</small></td>
        ${a.dates.map((s, i) => { if (i >= a.dim) return ''; const c = cellInfo(a, r, i + 1);
          return `<td class="${c.cls}" data-r="${r.row}" data-d="${i + 1}" title="${h(c.tip || '')}">${h(c.txt)}</td>`; }).join('')}
        <td class="sum m">${r.missing || ''}</td><td class="sum">${money(r.net)}</td></tr>`).join('')
        : `<tr><td colspan="40" class="empty">No employees in ${h(a.month)}.</td></tr>`}
      </tbody></table></div>`;
  $('#am').onchange = e => { if (Object.keys(S.pend).length && !confirm('Discard unsaved changes?')) { e.target.value = a.month; return; } S.month = e.target.value; viewAttendance(); };
  $('#aq').oninput = e => { S.attQ = e.target.value; const p = e.target.selectionStart; renderAttendance(); const n = $('#aq'); n.focus(); n.setSelectionRange(p, p); };
  $('#ar').onclick = () => { if (!Object.keys(S.pend).length || confirm('Discard unsaved changes?')) viewAttendance(true); };
  $('#ad').onclick = () => { S.pend = {}; renderAttendance(); };
  $('#as').onclick = saveAttendance;
  const mt = $('#mt'); if (mt) mt.onclick = () => {
    let n = 0;
    rows.forEach(r => { const c = cellInfo(a, r, todayIdx + 1); if (!c.ro && !c.txt) { S.pend[r.row + '|' + (todayIdx + 1)] = a.statuses[0].name; n++; } });
    toast(n ? `${n} employee(s) marked ${a.statuses[0].name} — click Save.` : 'Everyone is already marked today.'); renderAttendance();
  };
  document.querySelectorAll('table.att td.d').forEach(td => td.onclick = e => {
    const c = td.className.includes(' ro') || td.classList.contains('ro');
    selectCell(+td.dataset.r, +td.dataset.d);
    if (!c) showPop(td);
  });
  if (S.sel) markSel();
}
function selectCell(r, d) { S.sel = { r: r, d: d }; markSel(); }
function markSel() {
  document.querySelectorAll('table.att td.sel').forEach(x => x.classList.remove('sel'));
  const td = document.querySelector(`table.att td.d[data-r="${S.sel.r}"][data-d="${S.sel.d}"]`);
  if (td) { td.classList.add('sel'); td.scrollIntoView({ block: 'nearest', inline: 'nearest' }); }
}
function setCell(r, d, val) {
  const a = S.att, row = a.rows.find(x => x.row === r); if (!row || a.locked) return;
  const c = cellInfo(a, row, d); if (c.ro) return;
  const key = r + '|' + d;
  if ((row.att[d - 1] || '') === val) delete S.pend[key]; else S.pend[key] = val;
  hidePop(); renderAttendance();
}
function showPop(td) {
  const a = S.att, pop = $('#pop'), r = +td.dataset.r, d = +td.dataset.d;
  pop.innerHTML = a.statuses.map((s, i) => `<button data-v="${h(s.name)}"><span class="sw" style="background:var(--s${i});border:1px solid var(--t${i})"></span>${h(s.name)}<kbd>${i + 1}</kbd></button>`).join('') +
    `<button data-v=""><span class="sw" style="border:1px dashed var(--muted)"></span>Clear<kbd>Del</kbd></button>`;
  const b = td.getBoundingClientRect();
  pop.style.display = 'block';
  pop.style.left = Math.min(b.left, window.innerWidth - 180) + 'px';
  pop.style.top = (b.bottom + 4 + 190 > window.innerHeight ? b.top - 194 : b.bottom + 4) + 'px';
  pop.querySelectorAll('button').forEach(x => x.onclick = ev => { ev.stopPropagation(); setCell(r, d, x.dataset.v); });
}
function hidePop() { const p = $('#pop'); if (p) p.style.display = 'none'; }
document.addEventListener('click', e => { if (!e.target.closest('#pop') && !e.target.closest('td.d')) hidePop(); });
document.addEventListener('keydown', e => {
  if (S.view !== 'attendance' || !S.att || !S.sel || /input|select/i.test(document.activeElement.tagName)) return;
  const a = S.att, st = a.statuses;
  const rowsVis = Array.from(document.querySelectorAll('table.att tbody tr')).map(tr => { const td = tr.querySelector('td.d'); return td ? +td.dataset.r : null; }).filter(Boolean);
  let { r, d } = S.sel; const ri = rowsVis.indexOf(r);
  if (e.key === 'ArrowRight') d = Math.min(a.dim, d + 1); else if (e.key === 'ArrowLeft') d = Math.max(1, d - 1);
  else if (e.key === 'ArrowDown' && ri < rowsVis.length - 1) r = rowsVis[ri + 1]; else if (e.key === 'ArrowUp' && ri > 0) r = rowsVis[ri - 1];
  else if (e.key === 'Escape') { hidePop(); return; }
  else if (e.key === 'Delete' || e.key === 'Backspace') { setCell(r, d, ''); e.preventDefault(); return; }
  else {
    let k = /^[1-5]$/.test(e.key) ? Number(e.key) - 1 : st.findIndex(s => codeOf(s.name)[0] === e.key.toUpperCase());
    if (k >= 0 && st[k]) { setCell(r, d, st[k].name); S.sel = { r: r, d: Math.min(a.dim, d + 1) }; markSel(); e.preventDefault(); }
    return;
  }
  e.preventDefault(); hidePop(); selectCell(r, d);
});
async function saveAttendance() {
  const changes = Object.keys(S.pend).map(k => { const [r, d] = k.split('|').map(Number); return { row: r, day: d, value: S.pend[k] }; });
  if (!changes.length) return;
  const btn = $('#as'); btn.disabled = true; btn.innerHTML = '<span class="spin"></span> Saving & recalculating…';
  try {
    const res = await call('saveAttendance', { month: S.att.month, changes: changes });
    S.att = res.attendance; S.pend = {}; delete S.cache.summary; delete S.cache['reg:' + S.att.month];
    toast(`✔ Saved ${res.saved} change(s). Payroll recalculated.`); renderAttendance();
  } catch (e) { toast(e.message, true); btn.disabled = false; btn.textContent = `Save (${changes.length})`; }
}

/* ------------------------------- register ------------------------------- */
async function viewRegister(force) {
  const m = defMonth(), k = 'reg:' + m;
  if (!S.cache[k] || force) {
    loading('Loading ' + m + ' payroll register…');
    try { S.cache[k] = await call('register', { month: m }); } catch (e) { return failed(e, () => viewRegister(true)); }
  }
  const d = S.cache[k]; S.cur = d.currency || '₹';
  const cols = [['Employee ID', r => r.id], ['Employee Name', r => r.name], ['Dept', r => r.dept], ['Shift', r => r.shift], ['Monthly Salary', r => r.rate, 1],
    ['Working Days', r => r.workDays, 2]].concat(d.statuses.map((s, i) => [s, r => r.counts[i], 2]))
    .concat([['Sandwich', r => r.sandwich, 2], ['Missing', r => r.missing, 2], ['Paid Days', r => r.paidDays, 2], ['Unpaid (LOP)', r => r.unpaid, 2],
      ['Earned', r => r.earned, 1], ['Other Earn.', r => r.otherEarn, 1], ['Gross', r => r.gross, 1], ['LOP Ded.', r => r.lop, 1], ['EPF Emp.', r => r.epfEe, 1],
      ['EPF Empr.', r => r.epfEr, 1], ['ESI', r => r.esi, 1], ['Other Ded.', r => r.otherDed, 1], ['Advance/Loan', r => r.advance, 1], ['Total Ded.', r => r.totalDed, 1],
      ['NET PAYABLE', r => r.net, 1], ['Remarks', r => r.remarks]]);
  const fmt = (c, v) => c[2] === 1 ? money(v) : c[2] === 2 ? num(v) : h(v);
  $('#view').innerHTML = `<div class="toolbar"><select id="rm">${monthOpts(d.month)}</select>
      ${d.locked ? '<span class="pill lock">🔒 Finalized</span>' : '<span class="pill ok"><span class="dot"></span>Live</span>'}
      <span class="hint">${d.rows.length} employee(s)</span><div class="grow"></div>
      <button class="btn" id="rr">↻ Refresh</button><button class="btn" id="rc">⬇ Excel (CSV)</button><button class="btn pri" id="rp">⬇ PDF</button></div>
    <div class="scroll"><table class="t"><thead><tr>${cols.map(c => `<th class="${c[2] ? 'n' : ''}">${h(c[0])}</th>`).join('')}</tr></thead><tbody>
      ${d.rows.map(r => `<tr>${cols.map((c, i) => { const v = c[1](r);
        const cls = c[0] === 'Remarks' ? 'rem' + (/^MISSING/.test(v) ? ' miss' : '') : c[2] ? 'n' : '';
        return `<td class="${cls}${c[0] === 'NET PAYABLE' ? ' net' : ''}">${fmt(c, v)}</td>`; }).join('')}</tr>`).join('')}
      </tbody><tfoot><tr>${cols.map((c, i) => `<td class="${c[2] ? 'n' : ''}">${i === 0 ? 'TOTAL' : (c[2] && c[0] !== 'Working Days' && c[0] !== 'Monthly Salary') ? fmt(c, d.rows.reduce((a, r) => a + (Number(c[1](r)) || 0), 0)) : ''}</td>`).join('')}</tr></tfoot></table></div>`;
  $('#rm').onchange = e => { S.month = e.target.value; viewRegister(); };
  $('#rr').onclick = () => viewRegister(true);
  $('#rc').onclick = async () => {
    const esc = v => { const s = String(v); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
    const csv = [cols.map(c => esc(c[0])).join(',')].concat(d.rows.map(r => cols.map(c => esc(c[1](r))).join(','))).join('\r\n');
    const res = await P.saveCsv(csv, `Payroll Register - ${d.month} ${d.year}.csv`); if (res.saved) toast('Saved ' + res.path);
  };
  $('#rp').onclick = () => pdfAction($('#rp'), 'registerPdf', { month: d.month });
}
async function pdfAction(btn, action, params) {
  const t = btn.innerHTML; btn.disabled = true; btn.innerHTML = '<span class="spin"></span> Creating PDF…';
  try { const r = await call(action, params); const s = await P.savePdf(r.base64, r.fileName); if (s.saved) toast('✔ Saved ' + s.path); return r; }
  catch (e) { toast(e.message, true); } finally { btn.disabled = false; btn.innerHTML = t; }
}

/* ------------------------------- salary slip ---------------------------- */
async function viewSlip(force) {
  if (!S.cache.emp || force) {
    loading();
    try { S.cache.emp = await call('employees'); } catch (e) { return failed(e, () => viewSlip(true)); }
  }
  const E = S.cache.emp; S.cur = E.currency || '₹';
  const m = defMonth(), id = S.slipId || (E.employees[0] && E.employees[0].id);
  const k = 'reg:' + m;
  if (!S.cache[k]) { try { S.cache[k] = await call('register', { month: m }); } catch (e) { /* preview optional */ } }
  const reg = S.cache[k], r = reg && reg.rows.find(x => x.id === id), emp = E.employees.find(x => x.id === id);
  const line = (l, v) => `<tr><td>${h(l)}</td><td class="n">${money(v)}</td></tr>`;
  $('#view').innerHTML = `<div class="toolbar">
      <select id="se" style="min-width:260px">${E.employees.map(e => `<option value="${h(e.id)}" ${e.id === id ? 'selected' : ''}>${h(e.id)} - ${h(e.name)}${e.status !== 'Active' ? ' (' + h(e.status) + ')' : ''}</option>`).join('')}</select>
      <select id="sm">${monthOpts(m)}</select>
      <label class="hint"><input type="checkbox" id="sd" ${S.cfg.slipToDrive ? 'checked' : ''}> Also save a copy in Drive ▸ Salary Slips</label>
      <div class="grow"></div>
      ${isAdmin() ? '<button class="btn" id="sx">✉ Email to employee</button>' : ''}
      <button class="btn pri" id="sp">📄 Generate Salary Slip PDF</button></div>
    ${!emp ? '<div class="empty"><b>No employees</b>Add employees in the Google Sheet ▸ Settings ▸ Employee Master.</div>' : `
    <div class="grid g2">
      <div class="card"><h3>Employee</h3><table class="t"><tbody>
        <tr><td>Employee</td><td><b>${h(emp.id)} — ${h(emp.name)}</b></td></tr><tr><td>Department / Designation</td><td>${h(emp.dept)} / ${h(emp.desig)}</td></tr>
        <tr><td>Shift</td><td>${h(r ? r.shift : emp.shift)}</td></tr><tr><td>Joining date</td><td>${h(emp.join)}</td></tr>
        <tr><td>Email</td><td>${h(emp.email) || '<span class="hint">not set</span>'}</td></tr><tr><td>Status</td><td>${h(emp.status)}${emp.exit ? ' · last day ' + h(emp.exit) : ''}</td></tr>
      </tbody></table></div>
      <div class="card"><h3>${h(m)} preview</h3>${!r ? '<p class="hint">No attendance for this month yet — the slip cannot be generated.</p>' : `
        <table class="t"><tbody>${line('Earned salary', r.earned)}${line('Other earnings', r.otherEarn)}<tr><td><b>Gross</b></td><td class="n"><b>${money(r.gross)}</b></td></tr>
        ${line('LOP / absence', r.lop)}${line('EPF – employee', r.epfEe)}${line('EPF – employer (CTC)', r.epfEr)}${line('ESI – employer (CTC)', r.esi)}
        ${line('Other deduction', r.otherDed)}${line('Advance / loan recovery', r.advance)}<tr><td><b>Total deductions</b></td><td class="n"><b>${money(r.totalDed)}</b></td></tr>
        <tr><td><b>NET SALARY</b></td><td class="n net" style="font-size:16px">${money(r.net)}</td></tr></tbody></table>
        <p class="hint">Paid days ${num(r.paidDays)} · Unpaid ${num(r.unpaid)}${r.missing ? ` · <b style="color:var(--t4)">${r.missing} missing</b>` : ''}. The PDF is the exact slip from the sheet.</p>`}</div>
    </div>`}`;
  $('#se').onchange = e => { S.slipId = e.target.value; viewSlip(); };
  $('#sm').onchange = e => { S.month = e.target.value; viewSlip(); };
  $('#sd').onchange = e => { S.cfg.slipToDrive = e.target.checked; P.setConfig({ slipToDrive: e.target.checked }); };
  $('#sp').onclick = () => pdfAction($('#sp'), 'slipPdf', { id: id, month: m, saveToDrive: !!S.cfg.slipToDrive });
  const sx = $('#sx'); if (sx) sx.onclick = async () => {
    if (!emp || !emp.email) return toast('No email for this employee in the Employee Master.', true);
    if (!confirm(`Send the ${m} salary slip of ${emp.name}\nto ${emp.email}?${r ? '\n\nNet salary: ' + money(r.net) : ''}`)) return;
    sx.disabled = true; sx.innerHTML = '<span class="spin"></span> Sending…';
    try { const x = await call('emailSlip', { id: id, month: m }); toast('✔ Slip emailed to ' + x.to); } catch (e) { toast(e.message, true); }
    sx.disabled = false; sx.innerHTML = '✉ Email to employee';
  };
}

/* ------------------------------- employees ------------------------------ */
async function viewEmployees(force) {
  if (!S.cache.emp || force) {
    loading();
    try { S.cache.emp = await call('employees'); } catch (e) { return failed(e, () => viewEmployees(true)); }
  }
  const E = S.cache.emp, q = (S.empQ || '').toLowerCase(); S.cur = E.currency || '₹';
  const list = E.employees.filter(e => !q || Object.values(e).join(' ').toLowerCase().includes(q));
  const stc = s => s === 'Active' ? 'ok' : (s === 'Inactive' ? '' : 'bad');
  $('#view').innerHTML = `<div class="toolbar"><input type="search" id="eq" placeholder="Search name, ID, department…" value="${h(S.empQ || '')}" style="width:280px">
      <span class="hint">${list.length} of ${E.employees.length}</span><div class="grow"></div>
      <button class="btn" id="er">↻ Refresh</button><button class="btn pri" id="ee">✎ Add / edit in Google Sheet</button></div>
    <div class="scroll"><table class="t"><thead><tr><th>ID</th><th>Name</th><th>Department</th><th>Designation</th><th>Shift</th><th>Joining</th>
      <th class="n">Monthly Salary</th>${E.heads.map(x => `<th class="n">${h(x)}</th>`).join('')}<th>EPF</th><th class="n">Est. In-hand</th><th>Status</th><th>Email</th></tr></thead><tbody>
      ${list.map(e => `<tr><td><b>${h(e.id)}</b></td><td>${h(e.name)}</td><td>${h(e.dept)}</td><td>${h(e.desig)}</td><td>${h(e.shift)}</td><td>${h(e.join)}</td>
        <td class="n"><b>${money(e.salary)}</b></td>${e.heads.map(x => `<td class="n">${x.amount ? money(x.amount) : ''}</td>`).join('')}<td>${h(e.epf)}</td>
        <td class="n">${money(e.inHand)}</td><td><span class="pill ${stc(e.status)}">${h(e.status)}</span>${e.exit ? ' <span class="hint">' + h(e.exit) + '</span>' : ''}</td><td>${h(e.email)}</td></tr>`).join('')}
      </tbody></table></div>
    <p class="hint">Employees, salary structure, shift changes, increments and advances are maintained in the Google Sheet ▸ Settings tab, exactly as before.</p>`;
  $('#eq').oninput = e => { S.empQ = e.target.value; const p = e.target.selectionStart; viewEmployees(); const n = $('#eq'); n.focus(); n.setSelectionRange(p, p); };
  $('#er').onclick = () => viewEmployees(true);
  $('#ee').onclick = () => openSheet(false);
}

/* ------------------------------- sheet ---------------------------------- */
async function openSheet(inBrowser) { const r = await P.openSheet(inBrowser); if (!r.ok) toast(r.error, true); }
function viewSheet() {
  $('#view').innerHTML = `<div class="grid g2"><div class="card"><h3>Your live Google Sheet</h3>
    <p>Everything from the original system works here unchanged — the <b>Payroll System</b> menu, buttons, Settings tab, Employee Master,
       holidays, shift changes, increments, advances, finalize/lock, PDFs and email.</p>
    <div class="toolbar"><button class="btn pri" id="o1">⊞ Open inside the app</button><button class="btn" id="o2">↗ Open in my browser</button></div>
    <p class="hint">First time inside the app: sign in with the Google account that has access to the sheet. If Google says the browser is not secure, use "Open in my browser".</p></div>
    <div class="card"><h3>What to do where</h3><ul class="steps">
      <li><b>Desktop app:</b> daily attendance, payroll register, salary slip PDFs & email, dashboard.</li>
      <li><b>Google Sheet:</b> add / edit employees and salary heads, holidays, weekly offs, shift changes, increments, advances & loans, exits, company settings.</li>
      <li>Both always show the same data — changes appear in the app after ↻ Refresh.</li></ul></div></div>`;
  $('#o1').onclick = () => openSheet(false); $('#o2').onclick = () => openSheet(true);
}

/* ------------------------------- admin tools ---------------------------- */
async function viewTools() {
  if (!isAdmin()) { $('#view').innerHTML = '<div class="empty"><b>Admin only</b>Connect with the ADMIN key to use these tools.</div>'; return; }
  const sf = await P.scriptFiles();
  const m = defMonth();
  $('#view').innerHTML = `<div class="grid g2">
    <div class="card"><h3>Payroll engine</h3>
      <div class="toolbar"><button class="btn pri" data-a="recalc">🧮 Recalculate all months</button><button class="btn" data-a="sync">🔄 Sync employees</button>
        <button class="btn" data-a="markOffs">📅 Mark weekly offs & holidays</button></div>
      <div class="toolbar"><button class="btn" data-a="selfTest">✅ Run self-test</button><span class="hint">Takes 1–3 minutes.</span></div>
      <div id="tout" class="hint"></div></div>
    <div class="card"><h3>Finalize / unlock a month</h3>
      <div class="toolbar"><select id="lm">${monthOpts(m)}</select><button class="btn red" id="lf">🔒 Finalize</button><button class="btn" id="lu">🔓 Unlock</button></div>
      <p class="hint">Finalize freezes a paid month so later changes to salary, EPF or shifts never alter it. Unlock recalculates it from current data.</p></div>
    <div class="card"><h3>Sheet script</h3>
      <p>Google Sheet runs script <b>v${h(S.ping.scriptVersion)}</b> (API ${h(S.ping.api)}). This app includes <b>v${h(sf.scriptVersion)}</b> (API ${h(sf.apiVersion)}).
        ${sf.scriptVersion === S.ping.scriptVersion && sf.apiVersion === S.ping.api ? '<span class="pill ok">Up to date</span>' : '<span class="pill bad">Update available</span>'}</p>
      <div class="toolbar">${sf.files.map(f => `<button class="btn" data-copy="${h(f.name)}">📋 Copy ${h(f.name)} (${f.lines} lines)</button>`).join('')}</div>
      <ol class="steps hint"><li>Open the Google Sheet ▸ <b>Extensions ▸ Apps Script</b>.</li><li>Open the file with the same name, select all, paste, <b>Save</b>.</li>
        <li><b>Deploy ▸ Manage deployments ▸ ✏ Edit ▸ Version: New version ▸ Deploy</b> (same URL — nobody needs to reconnect).</li>
        <li>Reload the sheet; if the new version says so, run Payroll System ▸ Upgrade.</li></ol></div>
    <div class="card"><h3>Activity log</h3><p class="hint">Every change made from the desktop app (attendance saves, PDFs, emails, recalculations, locks) is recorded with the user's name in the hidden <b>App Log</b> tab of the sheet.</p></div>
  </div>`;
  document.querySelectorAll('[data-a]').forEach(b => b.onclick = async () => {
    const t = b.innerHTML; b.disabled = true; b.innerHTML = '<span class="spin"></span> Working…';
    try {
      const r = await call(b.dataset.a);
      $('#tout').innerHTML = Array.isArray(r) ? `<ul class="res" style="padding:0">${r.map(x => `<li>${h(x)}</li>`).join('')}</ul>` : '✔ ' + h(r);
      S.cache = {}; S.att = null;
    } catch (e) { toast(e.message, true); }
    b.disabled = false; b.innerHTML = t;
  });
  const lk = async (action, b) => {
    const mm = $('#lm').value;
    if (action === 'finalize' && !confirm(`Lock ${mm}? Its payroll figures will be frozen.`)) return;
    b.disabled = true; try { toast(await call(action, { month: mm })); S.cache = {}; S.att = null; } catch (e) { toast(e.message, true); } b.disabled = false;
  };
  $('#lf').onclick = () => lk('finalize', $('#lf')); $('#lu').onclick = () => lk('unlock', $('#lu'));
  document.querySelectorAll('[data-copy]').forEach(b => b.onclick = async () => { if (await P.copyScript(b.dataset.copy)) toast(b.dataset.copy + ' copied — paste it in Apps Script.'); });
}

/* ------------------------------- settings ------------------------------- */
function viewSettings() {
  const c = S.cfg;
  $('#view').innerHTML = `<div class="grid g2">
    <div class="card"><h3>Connection to Google Sheet</h3>
      <label class="f">Your name (shown in the activity log)</label><input type="text" id="cn" value="${h(c.userName || '')}" placeholder="e.g. Priya – HR">
      <label class="f">Web App URL (from Apps Script ▸ Deploy)</label><input type="text" id="cu" value="${h(c.webAppUrl || '')}" placeholder="https://script.google.com/macros/s/…/exec">
      <label class="f">Access key (ADMIN or STAFF)</label><input type="password" id="ck" value="${h(c.key || '')}" placeholder="ADM-… or STF-…">
      <div class="toolbar" style="margin-top:14px"><button class="btn pri" id="ct">Test & Save</button><span id="cr" class="hint"></span></div>
      <p class="hint">Get the URL and key from your administrator. Admin: run <code>setupDesktopApp</code> in Apps Script once (README, Part A).</p></div>
    <div class="card"><h3>App updates</h3>
      <p>Installed version <b>${h(c.appVersion)}</b> (${c.platform === 'darwin' ? 'Mac' : c.platform === 'win32' ? 'Windows' : h(c.platform)}).</p>
      <p id="updState">${updText(S.upd)}</p>
      <div class="toolbar"><button class="btn" id="cu2">Check for updates now</button></div>
      <p class="hint">The app checks automatically at start-up and every 4 hours. Updates never touch your payroll data — it lives in the Google Sheet.</p>
      <h3 style="margin-top:20px">Appearance</h3>
      <select id="th"><option value="">Follow system</option><option value="light" ${c.theme === 'light' ? 'selected' : ''}>Light</option><option value="dark" ${c.theme === 'dark' ? 'selected' : ''}>Dark</option></select></div></div>`;
  $('#ct').onclick = async () => {
    const conn = { webAppUrl: $('#cu').value.trim(), key: $('#ck').value.trim(), userName: $('#cn').value.trim() || 'User' };
    $('#cr').innerHTML = '<span class="spin"></span> Testing…';
    const r = await P.testConnection(conn);
    if (!r || !r.ok) { $('#cr').innerHTML = `<span style="color:var(--t4)">✕ ${h(r && r.error)}</span>`; return; }
    S.cfg = await P.setConfig(Object.assign(conn, { sheetUrl: r.data.spreadsheetUrl }));
    S.cfg = await P.getConfig(); S.cache = {}; S.att = null;
    $('#cr').innerHTML = `<span style="color:var(--t0)">✔ Connected to “${h(r.data.spreadsheetName)}” as ${r.data.role === 'admin' ? 'ADMIN' : 'STAFF'}</span>`;
    try { await connect(); setTimeout(() => go('dashboard'), 900); } catch (e) { toast(e.message, true); }
  };
  $('#cu2').onclick = () => { $('#updState').innerHTML = updText({ state: 'checking' }); P.checkUpdates(); };
  $('#th').onchange = e => { const v = e.target.value; if (v) document.documentElement.dataset.theme = v; else delete document.documentElement.dataset.theme; P.setConfig({ theme: v }); };
}

boot();
