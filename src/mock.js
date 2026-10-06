/* Demo data for trying the app without a Google Sheet:  PAYROLL_DEMO=1 npm start
 * Mirrors the shape of every DesktopApi.gs reply. Not used in normal operation.            */
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const ST = [{ name: 'Present', pct: 1 }, { name: 'WFH', pct: 0.5 }, { name: 'Half Day', pct: 0.5 }, { name: 'Leave', pct: 1 }, { name: 'Absent', pct: 0 }];
const now = new Date(), Y = now.getFullYear(), CM = now.getMonth();
const serial = (y, m, d) => Math.round((Date.UTC(y, m, d) - Date.UTC(1899, 11, 30)) / 86400000);
const today = serial(Y, CM, now.getDate());
const EMP = [['EMP001', 'Employee 1', 'Accounts', 'Accountant', 'Day', 30000], ['EMP002', 'Employee 2', 'Operations', 'Executive', 'Night', 22000],
  ['EMP003', 'Employee 3', 'Support', 'Night Shift Analyst', 'Night', 37000], ['EMP004', 'Employee 4', 'Support', 'Associate', 'Night', 25000],
  ['EMP005', 'Employee 5', 'Accounts', 'Trainee', 'Day (5-day)', 18000]];
const att = {};
function monthData(mi) {
  if (att[mi]) return att[mi];
  const dim = new Date(Y, mi + 1, 0).getDate();
  const rows = EMP.map((e, i) => {
    const a = [], t = [];
    for (let d = 1; d <= 31; d++) {
      if (d > dim) { a.push(''); t.push(''); continue; }
      const wd = new Date(Y, mi, d).getDay(), off = wd === 0 || (e[4] !== 'Day' && wd === 6);
      const ty = off ? 'Off' : 'Work', s = serial(Y, mi, d); t.push(ty);
      const h = (d * 7 + i * 13 + mi * 5) % 100;
      a.push(off ? 'Weekly Off' : s > today || (s === today && i > 2) ? '' : h < 3 ? 'Absent' : h < 7 ? 'Leave' : h < 12 ? 'Half Day' : h < 20 ? 'WFH' : 'Present');
    }
    return { row: 14 + i, id: e[0], name: e[1], shift: e[4], att: a, types: t, join: 0, exit: 0, other: '', missing: 0, net: 0, remarks: 'OK' };
  });
  att[mi] = { dim: dim, rows: rows }; recalc(mi); return att[mi];
}
function recalc(mi) {
  const m = att[mi];
  m.rows.forEach((r, i) => {
    let lop = 0, marked = 0, cnt = [0, 0, 0, 0, 0];
    r.att.forEach(v => { const k = ST.findIndex(s => s.name === v); if (k >= 0) { cnt[k]++; marked++; lop += 1 - ST[k].pct; } });
    const sal = EMP[i][5], ded = Math.round(lop * sal / 30), epf = i % 2 ? 0 : 3600;
    r.counts = cnt; r.marked = marked; r.lop = marked ? ded : 0; r.epf = marked ? epf : 0;
    r.gross = marked ? sal : 0; r.net = r.gross - r.lop - r.epf; r.remarks = marked ? 'OK' : 'Attendance not entered for this period';
  });
}
function register(mi) {
  const m = monthData(mi);
  return m.rows.filter(r => r.marked).map((r, i) => ({ id: r.id, name: r.name, dept: EMP[i][2], shift: r.shift, rate: EMP[i][5], workDays: 26, counts: r.counts,
    sandwich: 0, missing: 0, paidDays: 30 - r.lop * 30 / EMP[i][5], unpaid: Math.round(r.lop * 30 / EMP[i][5] * 100) / 100, earned: r.gross, otherEarn: 0,
    gross: r.gross, lop: r.lop, epfEe: r.epf / 2, epfEr: r.epf / 2, esi: 0, otherDed: 0, advance: 0, totalDed: r.lop + r.epf, net: r.net, remarks: r.remarks }));
}
const pdf = 'JVBERi0xLjQKMSAwIG9iago8PC9UeXBlL0NhdGFsb2cvUGFnZXMgMiAwIFI+PgplbmRvYmoKMiAwIG9iago8PC9UeXBlL1BhZ2VzL0tpZHNbMyAwIFJdL0NvdW50IDE+PgplbmRvYmoKMyAwIG9iago8PC9UeXBlL1BhZ2UvUGFyZW50IDIgMCBSL01lZGlhQm94WzAgMCAyMDAgMjAwXT4+CmVuZG9iagp0cmFpbGVyCjw8L1Jvb3QgMSAwIFI+PgolJUVPRgo=';
module.exports = function mock(action, params) {
  const mi = params && params.month ? MONTHS.indexOf(params.month) : CM;
  const D = (data) => ({ ok: true, data: data });
  switch (action) {
    case 'ping': return D({ role: 'admin', api: '1.0.0', scriptVersion: '2.3', setup: '', spreadsheetName: 'Payroll (demo)', spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/demo', company: 'Demo Company Pvt. Ltd.', year: Y, capacity: 100 });
    case 'summary': return D({ company: 'Demo Company Pvt. Ltd.', year: Y, currency: '₹', currentMonth: MONTHS[CM], statuses: ST, shifts: ['Day', 'Night', 'Day (5-day)'],
      months: MONTHS.map((m, i) => { const r = i <= CM ? register(i) : []; const s = k => r.reduce((a, x) => a + x[k], 0);
        const c = [0, 1, 2, 3, 4].map(k => r.reduce((a, x) => a + x.counts[k], 0)), tot = c.reduce((a, b) => a + b, 0);
        return { month: m, locked: i < CM - 1, active: r.length, counts: c, attPct: tot ? c[0] / tot : null, paidPct: tot ? (c[0] + c[1] * .5 + c[2] * .5 + c[3]) / tot : null,
          missing: 0, gross: s('gross'), absence: s('lop'), epf: s('epfEe') + s('epfEr'), other: 0, net: s('net') }; }),
      today: { date: now.toDateString(), counts: [2, 1, 0, 0, 0], missing: 2, expected: 5 }, employees: { total: 5, active: 5, inactive: 0, exited: 0, byShift: {} } });
    case 'employees': return D({ currency: '₹', heads: ['Basic', 'Dearness Allowance', 'Special Allowance'], employees: EMP.map((e, i) => ({ id: e[0], name: e[1], dept: e[2], desig: e[3],
      shift: e[4], join: '01-Apr-2024', salary: e[5], email: e[1].replace(' ', '').toLowerCase() + '@example.com', status: 'Active', epf: i % 2 ? 'No' : 'Yes', inHand: e[5] - (i % 2 ? 0 : 3600),
      heads: [{ name: 'Basic', amount: e[5] / 2 }, { name: 'Dearness Allowance', amount: e[5] / 3 }, { name: 'Special Allowance', amount: e[5] / 6 }], exit: '' })) });
    case 'attendance': case 'saveAttendance': {
      const m = monthData(mi);
      if (action === 'saveAttendance') { (params.changes || []).forEach(c => { const r = m.rows.find(x => x.row === c.row); if (r) r.att[c.day - 1] = c.value; }); recalc(mi); }
      const out = { month: MONTHS[mi], year: Y, locked: mi < CM - 1, dim: m.dim, statuses: ST, currency: '₹', today: today, trackStart: serial(Y, 0, 1),
        dates: Array.from({ length: 31 }, (_, d) => d < m.dim ? serial(Y, mi, d + 1) : 0),
        dayNames: Array.from({ length: 31 }, (_, d) => d < m.dim ? new Date(Y, mi, d + 1).toDateString().slice(0, 3) : ''), rows: m.rows };
      return action === 'saveAttendance' ? D({ saved: params.changes.length, attendance: out }) : D(out);
    }
    case 'register': return D({ month: MONTHS[mi], year: Y, locked: mi < CM - 1, statuses: ST.map(s => s.name), currency: '₹', company: 'Demo', rows: register(mi) });
    case 'slipPdf': case 'registerPdf': return D({ fileName: 'demo.pdf', base64: pdf });
    case 'selfTest': return D(['✅ Demo self-test passed']);
    default: return D('Done (demo).');
  }
};
