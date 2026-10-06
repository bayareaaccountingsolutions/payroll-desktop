/******************************************************************************
 *  EMPLOYEE ATTENDANCE + PAYROLL + SALARY SLIP SYSTEM  —  Google Sheets
 *  Version 2.3  (employee capacity set from the menu: Payroll System ▸ Change employee capacity — keeps all data;
 *                v2.2 hardening included: bad rows show a remark instead of stopping the script)
 *  EXISTING FILE?  Paste this code, Save, reload the sheet, then Payroll System ▸ ⬆️ Upgrade to v2 (keeps all data).
 *
 *  ONE-TIME INSTALL (about 3 minutes)
 *   1. Open a NEW blank Google Sheet  (type  sheets.new  in the browser).
 *   2. Menu:  Extensions ▸ Apps Script.  Delete everything in Code.gs,
 *      paste this ENTIRE file, click Save (disk icon).
 *   3. In the toolbar function list choose  setupPayrollSystem  ▸ click Run.
 *      Approve the permission prompts (Advanced ▸ Go to project ▸ Allow).
 *   4. Go back to the spreadsheet tab. Setup runs in stages (about 5–10 min in
 *      total). If it needs more than one run it CONTINUES AUTOMATICALLY — the
 *      progress shows on the Dashboard tab and in pop-up messages. When you see
 *      "Setup complete", reload the browser tab once — a new menu
 *      "Payroll System" appears.
 *   Capacity: 100 employees for a new file. To grow: Payroll System ▸ Change employee capacity (no code edits).
 *
 *  SPEED: all payroll maths runs in this script and is written to the sheets
 *  as plain values, so the file opens and responds quickly. It recalculates
 *  automatically after every edit (a few seconds). If figures ever stop
 *  updating, use Payroll System ▸ Enable Automatic Calculation, or
 *  Payroll System ▸ Recalculate Payroll Now.
 *
 *  Tabs created (in this order):
 *   Dashboard | Salary Slip | Payroll Register | January … December | Settings
 *
 *  Daily use:  Settings ▸ Employee Master (add people)  →  month tab (pick
 *  attendance from dropdowns)  →  Salary Slip (pick employee + month)  →
 *  "Generate Salary Slip PDF" button.  Everything else is automatic.
 *
 *  PROTECTED RANGES (warning-only: editors get a warning before overwriting)
 *   • Month sheets: everything EXCEPT the day cells (H:AL) and Other Earnings (AY)
 *   • Dashboard:    everything EXCEPT the month selector N15
 *   • Salary Slip:  everything EXCEPT Select Employee C2 and Select Month F2
 *   • Settings:     only the auto cells — Daily Salary I24:I73, shift auto
 *                   columns I5:J8, weekly-off headings G11:J11, active list
 *                   AF5:AF73, auto lists AH5:AI16, Current Payroll Month C14,
 *                   Monthly Gross G24:G73, Est. In-hand Z24:Z73,
 *                   Current Shift/Salary P24:Q73 and the auto columns of G & H
 *   To make them strict: Data ▸ Protect sheets and ranges ▸ choose one ▸
 *   "Restrict who can edit this range".
 *
 *  KEY RULES
 *   • Night shift = ONE entry on the date the shift starts.
 *   • Weekly offs are set per shift (default: Day = Sunday; Night = Sat + Sun;
 *     "Day (5-day)" = Sat + Sun for day-shift staff with a 5-day week).
 *   • Weekly offs, holidays, pre-joining and future dates: never Absent,
 *     never Missing, never deducted. Mid-month joiners are paid pro-rata.
 *   • WFH and Half Day are paid at 50% (editable in Settings ▸ C).
 *   • Net Salary = (Earned Salary + Other Earnings) − (Absence + EPF + ESI + Other).
 *   • Salary structure: Basic, DA, allowances and ESI Employer per employee in
 *     the Employee Master (columns R–Y); Monthly Gross = sum of heads.
 *   • Sandwich leave: weekly offs / holidays between two Leave/Absent days are
 *     deducted too (Settings C18/C19, status table column N).
 *   • EPF is flat: No → 0; Yes → employee share ₹1,800 (+ employer share ₹1,800
 *     when Settings C17 = Yes, i.e. salary is CTC → ₹3,600 total from salary);
 *     blank amounts use the Settings defaults.
 *   • Weekly offs / holidays are written into day cells as "Weekly Off" / "Holiday".
 *   • Shift moves and salary increments are entered with an EFFECTIVE DATE in
 *     Settings sections G and H — never overwrite the base shift / salary.
 *     Mid-month changes are applied day by day (weekly offs follow the shift
 *     in force that day; salary is pro-rated old rate / new rate).
 *   • Payroll System ▸ Finalize (lock) a month freezes a paid month so later
 *     changes (salary, EPF, shift) never alter it.
 *   • ADVANCES / LOANS (Settings ▸ I, row 180): One-time = whole balance from the
 *     "Recover From" month's net salary; Monthly = instalment every month until the
 *     balance is zero. Shortfalls carry forward; oldest advance recovered first;
 *     never more than Settings C20 (% of net salary). Recorded month by month.
 *   • EXITS: Employment Status = Terminated / Discontinued + Last Working Day
 *     (Employee Master column AA) → paid pro-rata up to that day, removed from all
 *     later months, any advance balance recovered from the final salary.
 ******************************************************************************/

/* ============================== CONFIGURATION ============================== */
const APP = {
  ROWS: 50,               // employee capacity — set from the menu (Payroll System ▸ Change employee capacity), never edit here
  MASTER_FIRST: 24,       // first Employee Master data row on Settings
  M_FIRST: 14,            // first employee row on each month sheet
  HOL_ROWS: 60,           // holiday list capacity
  TAG: 'PayrollSystem',   // marks protections created by this script
  // Brand colours (edit to taste)
  NAVY: '#1F2A44', NAVY2: '#2E3B5E', RED: '#E4141A', RED_LIGHT: '#FDECEC',
  PALE: '#EEF1F7', INPUT: '#FFF8E1', AUTO: '#F3F4F6', BORDER: '#C9CFDB',
  TEXT: '#1F2937', MUTED: '#6B7280', WHITE: '#FFFFFF', GREEN: '#1E7B45'
};
APP.CHG_ROWS = 50; APP.ADV_ROWS = 50;
/** v2.3: every row position follows the employee capacity stored in this file (no code edits needed to grow).
 *  Files without a stored capacity keep the original 50-employee layout exactly. */
function layoutFor_(cap, legacy) {
  const L = { ROWS: cap };
  L.MASTER_LAST = APP.MASTER_FIRST + cap - 1;                 // Employee Master rows 24 … MASTER_LAST
  L.M_LAST = APP.M_FIRST + cap - 1;                           // month-tab employee rows 14 … M_LAST
  L.CHG_TITLE = L.MASTER_LAST + (legacy ? 52 : 4);            // G. Shift changes / H. Increments (title row)
  L.CHG_FIRST = L.CHG_TITLE + 3; L.CHG_LAST = L.CHG_FIRST + APP.CHG_ROWS - 1;
  L.ADV_TITLE = L.CHG_LAST + 3;                               // I. Advance & Loan Register (title row)
  L.ADV_FIRST = L.ADV_TITLE + 3; L.ADV_LAST = L.ADV_FIRST + APP.ADV_ROWS - 1;
  return L;
}
function storedLayout_() {
  try { const v = PropertiesService.getDocumentProperties().getProperty('CAPACITY'); if (v && Number(v) >= 10) return { cap: Number(v), legacy: false }; } catch (e) { }
  return { cap: 50, legacy: true };
}
function applyLayout_(cap, legacy) { Object.assign(APP, layoutFor_(cap, legacy)); APP.LEGACY = legacy; }
(function () { const s = storedLayout_(); applyLayout_(s.cap, s.legacy); })();
APP.EXIT_COL = 27;                                                     // Employee Master column AA = Last Working Day
APP.HOL_COL = 28;                                     // Holiday calendar in Settings columns AB:AD
APP.HEADS = 7;                                        // salary heads in Employee Master R..X (ESI Employer in Y)
APP.TIME_BUDGET_MS = 4 * 60 * 1000;                   // setup works in chunks of ≤4 min (Google limit is 6)

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
  'August', 'September', 'October', 'November', 'December'];
const M30 = 'Monthly Salary ÷ 30';
const MWD = 'Monthly Salary ÷ Working Days';
const MDEF = 'Company Default';
const OFF_LABEL = 'Weekly Off', HOL_LABEL = 'Holiday';   // auto-filled into day cells
const DAYNAMES = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

// Attendance-cell colours: [background, font] for Present, WFH, Half Day, Leave, Absent
const STATUS_COLORS = [
  ['#D9F2E3', '#1E7B45'], ['#DCEBFB', '#1A5FB4'], ['#FFF1CC', '#8A5A00'],
  ['#EADCF8', '#6B3FA0'], ['#F9D6D5', '#B3261E']
];
const SANDWICH_BG = '#FDBA74';
const MISSING_BG = '#FF8A80', OFF_BG = '#E3E6EA', HOL_BG = '#D5EFEC', PREJOIN_BG = '#F1F1F1', NOMONTH_BG = '#B8BEC8';

/* Month-sheet column map (fixed layout) */
const MC = {
  ID: 'A', NAME: 'B', DEPT: 'C', DESIG: 'D', SHIFT: 'E', SAL: 'F', DAILY: 'G',
  D1: 'H', D31: 'AL',
  PRESENT: 'AM', WFH: 'AN', HALF: 'AO', LEAVE: 'AP', ABSENT: 'AQ', MISSING: 'AR',
  ELIGWD_SHOW: 'AS', PAIDDAYS: 'AT', UNPAID: 'AU', ATTPCT: 'AV', PAIDPCT: 'AW',
  EARNED: 'AX', OTHEREARN: 'AY', GROSS: 'AZ', ABSDED: 'BA', EPS: 'BB', OTHERDED: 'BC',
  TOTDED: 'BD', NET: 'BE', REMARKS: 'BF',
  // hidden helpers
  JOIN: 'BG', METHOD: 'BH', STATUS: 'BI', ELIGCAL: 'BJ', ELIGWD: 'BK', MARKED: 'BL',
  PAIDW: 'BM', SHIFTROW: 'BN', SHIFTWD: 'BO', TODAYEXP: 'BP', TODAYENTRY: 'BQ', EPF_EE: 'BR', EPF_ER: 'BS',
  SHIFTCHG: 'BT', INCCHG: 'BU', TYPE1: 'BV', TYPE31: 'CZ', RATE1: 'DA', RATE31: 'EE',  // per-day shift type & salary rate
  ESI: 'EF', SWC: 'EG', SWM: 'EH', BRW: 'EI', BRM: 'EJ'   // ESI, sandwich days/money, sandwich-border days/money
};
// Sandwich-day flags (hidden, 1 = weekly off / holiday deducted as sandwich) in EK..FO
const SWB = { FLAG1: 141, LAST: 171 };   // column numbers (EK .. FO)
// v2: advance / loan recovery & exit columns, placed after the hidden helper block
//  FP = Advance / Loan Recovery (VISIBLE — shows right after Remarks), FQ = net before advance, FR = last working day,
//  FS = advance balance outstanding after the month (FQ..FS hidden)
const XC = { ADV: 172, NETPRE: 173, EXIT: 174, ADVBAL: 175 };
const OUT_LAST = 175;
const EXITED = ['Terminated', 'Discontinued'];
// Columns written by the calculation engine (values) — everything else on a month sheet is input or a light formula
const OUT_FIRST = 39;                      // AM


const BTN_PDF_B64 = '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';
const BTN_MAIL_B64 = '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';
const BTN_REG_B64 = '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';
const REG = 'Payroll Register';   // tab name

/* ================================ MENUS ==================================== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Payroll System')
    .addItem('📄 Generate Salary Slip PDF', 'generateSalarySlipPDF')
    .addItem('✉️ Send Salary Slip by Email', 'sendSalarySlipEmail')
    .addItem('📊 Download Payroll Register PDF', 'downloadRegisterPDF')
    .addSeparator()
    .addItem('🧮 Recalculate Payroll Now', 'menuRecalc')
    .addItem('⚡ Enable Automatic Calculation', 'menuAuto')
    .addItem('🔄 Sync Employees to Month Sheets', 'menuSync')
    .addItem('📅 Mark Weekly Offs & Holidays', 'menuMarkOffs')
    .addItem('🔒 Finalize (lock) a month', 'menuFinalize')
    .addItem('🔓 Unlock a month', 'menuUnlock')
    .addItem('₹ Apply Currency Format', 'menuCurrency')
    .addItem('🎨 Refresh Attendance Colours', 'menuColours')
    .addSeparator()
    .addItem('📒 Advance & Loan Register', 'goAdvanceRegister')
    .addItem('🚪 Exit an employee (Terminate / Discontinue)', 'goEmployeeMaster')
    .addItem('⬆️ Upgrade to v2 (keeps your data)', 'menuUpgrade')
    .addItem('📈 Change employee capacity', 'changeCapacity')
    .addSeparator()
    .addItem('✅ Run Self-Test', 'menuSelfTest')
    .addItem('🧹 Clear Sample Data', 'menuClearSample')
    .addSeparator()
    .addItem('⚠️ Rebuild Entire System', 'setupPayrollSystem')
    .addToUi();
  try { const ss = SpreadsheetApp.getActive(), g = ss.getSheetByName(REG); if (g && String(g.getRange('AD3').getValue()) !== registerMonth_(ss)) renderRegister_(); } catch (e) { }
  try {
    const props = PropertiesService.getDocumentProperties();
    if (props.getProperty('V2') !== '1' && props.getProperty('SETUP_STEP') === null && SpreadsheetApp.getActive().getSheetByName('Settings'))
      toast_('New features are ready: run Payroll System ▸ ⬆️ Upgrade to v2 (keeps your data).', 'Payroll — upgrade needed', 30);
  } catch (e) { }
}

/* ============================== SETUP (MAIN) =============================== */
function setupPayrollSystem() {
  const ss = SpreadsheetApp.getActive();
  const ui = safeUi_();
  if (ss.getSheetByName('Settings') && ui) {
    const ok = ui.alert('Rebuild Payroll System?',
      'This deletes ALL tabs in this file and rebuilds the system with sample data.\n' +
      'Any attendance you entered will be lost. Continue?', ui.ButtonSet.YES_NO);
    if (ok !== ui.Button.YES) return;
  }
  runSetupSteps_(true);
}

/** Called automatically by a one-off timer when setup needs more than one run (Google's 6-minute limit). */
function continueSetup() { runSetupSteps_(false); }

function setupSteps_() {
  const ss = SpreadsheetApp.getActive();
  const steps = [];
  steps.push(['Creating tabs & Settings', () => {
    ss.setSpreadsheetLocale('en_IN');
    ss.setSpreadsheetTimeZone('Asia/Kolkata');
    const temp = ss.insertSheet('__temp__' + Date.now());
    ss.getSheets().forEach(s => { if (s.getName() !== temp.getName()) ss.deleteSheet(s); });
    ss.getNamedRanges().forEach(n => n.remove());
    const order = ['Dashboard', 'Salary Slip', REG].concat(MONTHS).concat(['Settings']);
    order.forEach((n, i) => ss.insertSheet(n, i));
    ss.deleteSheet(temp);
    buildSettings_(ss.getSheetByName('Settings'));
    writeSampleData_(ss);
  }]);
  MONTHS.forEach((m, i) => steps.push(['Building ' + m, () => { buildMonth_(ss.getSheetByName(m), i + 1); writeSampleMonth_(ss, i); }]));
  steps.push(['Building Salary Slip', () => buildSalarySlip_(ss.getSheetByName('Salary Slip'))]);
  steps.push(['Building Payroll Register', () => buildRegister_(ss.getSheetByName(REG))]);
  steps.push(['Building Dashboard', () => buildDashboard_(ss.getSheetByName('Dashboard'))]);
  steps.push(['Calculating payroll', () => { recalcAll_(); renderRegister_(); installTriggers_(); }]);
  steps.push(['Colours, formats & protection', () => {
    refreshAttendanceFormatting();
    applyCurrencyFormats();
    applyProtections_(ss);
    ss.getSheetByName('Dashboard').activate();
  }]);
  return steps;
}

function runSetupSteps_(fresh) {
  const t0 = Date.now();
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getDocumentProperties();
  const ui = safeUi_();
  // only ONE setup may run at a time (a second run would delete tabs the first one is building)
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    if (fresh) { if (ui) ui.alert('Setup is already running', 'Another setup run is still working. Please wait for "Setup complete" — progress shows at the top of the Dashboard tab.', ui.ButtonSet.OK); }
    else { clearSetupTriggers_(); ScriptApp.newTrigger('continueSetup').timeBased().after(60 * 1000).create(); }
    return;
  }
  try {
    if (fresh) { clearSetupTriggers_(); props.setProperty('SETUP_STEP', '0'); props.deleteProperty('SETUP_ERROR'); props.setProperty('V2', '1');
      props.setProperty('CAPACITY', String(Math.max(100, APP.LEGACY ? 100 : APP.ROWS))); const sl = storedLayout_(); applyLayout_(sl.cap, sl.legacy); }
    const steps = setupSteps_();
    let step = Number(props.getProperty('SETUP_STEP') || 0), doneThisRun = 0;
    if (step > 0) defineNamedRanges_(ss);                                  // self-heal names on every resumed run
    while (step < steps.length) {
      if (doneThisRun > 0 && Date.now() - t0 > APP.TIME_BUDGET_MS) {   // every run finishes ≥1 stage; pause well before Google's 6-minute limit
        clearSetupTriggers_();
        ScriptApp.newTrigger('continueSetup').timeBased().after(20 * 1000).create();
        setupProgress_(ss, `⏳ Setup paused at step ${step + 1} of ${steps.length} — it continues AUTOMATICALLY in about 20 seconds. Keep this file open; do not edit yet.`);
        return;
      }
      setupProgress_(ss, `⏳ Setup in progress — step ${step + 1} of ${steps.length}: ${steps[step][0]}… (please wait, do not edit)`);
      try { steps[step][1](); }
      catch (err) {
        const msg = `Setup stopped at step ${step + 1} of ${steps.length} (${steps[step][0]}): ${err.message}`;
        props.setProperty('SETUP_ERROR', msg);
        setupProgress_(ss, '❌ ' + msg + ' — run setupPayrollSystem again.');
        if (ui) ui.alert('Setup could not finish', msg + '\n\nPlease run setupPayrollSystem again. If it stops at the same step, send this message to your administrator.', ui.ButtonSet.OK);
        throw err;
      }
      SpreadsheetApp.flush();
      step++; doneThisRun++;
      props.setProperty('SETUP_STEP', String(step));
    }
    defineNamedRanges_(ss);
    props.deleteProperty('SETUP_STEP');
    clearSetupTriggers_();
    toast_('Setup complete ✔  Reload the browser tab to see the "Payroll System" menu.', 'Done', 20);
    if (ui) ui.alert('Setup complete ✔',
      'The Attendance, Payroll & Salary Slip system is ready.\n\n' +
      '• Reload this browser tab once to see the "Payroll System" menu.\n' +
      '• Sample employees (EMP001–EMP005) and sample attendance are included — replace them via Settings, or use Payroll System ▸ Clear Sample Data.\n' +
      '• Run Payroll System ▸ Run Self-Test any time to verify every calculation.',
      ui.ButtonSet.OK);
  } finally { lock.releaseLock(); }
}

/** '' when the system is ready; otherwise a plain-language reason (setup still running / not set up). */
function setupState_() {
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getDocumentProperties();
  const err = props.getProperty('SETUP_ERROR');
  if (err) return err + '\n\nPlease run setupPayrollSystem again (Extensions ▸ Apps Script ▸ Run).';
  const step = props.getProperty('SETUP_STEP');
  if (step !== null) return `Setup is still running (finished ${step} of ${setupSteps_().length} steps). Please wait until you see "Setup complete" — progress shows at the top of the Dashboard tab.\n\nIf nothing changes for 10 minutes, run setupPayrollSystem again from Extensions ▸ Apps Script.`;
  if (!ss.getSheetByName('Settings') || MONTHS.some(m => !ss.getSheetByName(m))) return 'This file has not been set up yet. Open Extensions ▸ Apps Script, choose setupPayrollSystem and click Run.';
  if (props.getProperty('CAP_STEP') !== null) return 'The employee capacity is being changed. Please wait for "Capacity updated" (about 3–6 minutes). If nothing happens for 10 minutes, run Payroll System ▸ Change employee capacity again.';
  if (props.getProperty('V2') !== '1') return 'The script has been updated. Please run Payroll System ▸ ⬆️ Upgrade to v2 (keeps your data) once — it adds the Advance & Loan Register and exit handling.';
  return '';
}
/** Runs a menu/button action safely: waits for setup, repairs names, shows a clear message instead of a raw error. */
function safeRun_(fn, label) {
  const ui = safeUi_();
  const st = setupState_();
  if (st) { if (ui) ui.alert('Please wait', st, ui.ButtonSet.OK); return; }
  try { named_(SpreadsheetApp.getActive(), 'StatusNames'); fn(); }
  catch (err) {
    if (!ui) throw err;
    ui.alert(label + ' — could not complete', err.message + '\n\nTechnical detail (send this if it keeps happening):\n' +
      String(err.stack || '').split('\n').slice(0, 3).join('\n'), ui.ButtonSet.OK);
  }
}
// ----- menu entry points -----
function menuRecalc() { safeRun_(recalculatePayrollNow, 'Recalculate Payroll'); }
function menuAuto() { safeRun_(enableAutoCalculation, 'Enable Automatic Calculation'); }
function menuSync() { safeRun_(() => { syncEmployees(); markOffDays(); recalcAll_(); renderRegister_(); toast_('Employees synced ✔', 'Payroll', 5); }, 'Sync Employees'); }
function menuMarkOffs() { safeRun_(() => { markOffDays(); recalcAll_(); toast_('Weekly offs & holidays marked ✔', 'Payroll', 5); }, 'Mark Weekly Offs'); }
function menuFinalize() { safeRun_(finalizeMonth, 'Finalize month'); }
function menuUnlock() { safeRun_(unlockMonth, 'Unlock month'); }
function menuCurrency() { safeRun_(applyCurrencyFormats, 'Currency format'); }
function menuColours() { safeRun_(refreshAttendanceFormatting, 'Attendance colours'); }
function menuSelfTest() { safeRun_(runSelfTest, 'Self-Test'); }
function menuClearSample() { safeRun_(clearSampleData, 'Clear Sample Data'); }

function setupProgress_(ss, msg) {
  toast_(msg, 'Payroll setup', 30);
  const d = ss.getSheetByName('Dashboard');
  if (d && !d.getRange('X1').getValue()) d.getRange('A1').setValue(msg).setFontWeight('bold').setFontColor(APP.RED);   // until the Dashboard is built
}
function clearSetupTriggers_() {
  try { ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'continueSetup') ScriptApp.deleteTrigger(t); }); } catch (e) { }
}

/* ================================ SETTINGS ================================= */
function buildSettings_(sh) {
  sh.clear();
  ensureSize_(sh, APP.ADV_LAST + 5, 36);          // FIX v1.1: add the columns FIRST (a new sheet has only 26)
  sh.setTabColor(APP.NAVY);
  sh.setHiddenGridlines(true);
  const widths = [110, 170, 130, 140, 80, 110, 110, 190, 110, 95, 105, 105, 105, 220, 110, 115, 125,
    95, 95, 95, 95, 95, 95, 95, 95, 115, 20, 115, 210, 100, 20, 240, 20, 130, 130];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  banner_(sh.getRange('A1:AI1'), '⚙  SETTINGS — CONTROL PANEL', 16);
  sh.setRowHeight(1, 40);
  sh.getRange('A2:AI2').merge().setValue('Yellow cells are for you to edit. Grey cells calculate automatically. Every change flows to all month sheets, the Dashboard and the Salary Slip instantly.')
    .setFontColor(APP.MUTED).setFontStyle('italic');

  // ---------- A. Company settings (rows 4-19) ----------
  section_(sh.getRange('A3:D3'), 'A.  COMPANY SETTINGS');
  const year = Number(Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy'));
  const comp = [
    ['Company Name', 'Your Company Pvt. Ltd.'],
    ['Company Address', 'Office address, City, State – PIN'],
    ['Phone', '+91 00000 00000'],
    ['Email', 'hr@yourcompany.com'],
    ['Attendance Year  (controls all dates)', year],
    ['Currency Symbol', '₹'],
    ['Standard Working Hours / Day', 8],
    ['Salary Calculation Method', M30],
    ['Default EPF – Employee Share (flat ₹/month)', 1800],
    ['Attendance Tracking Start Date', '=DATE(C8,1,1)'],
    ['Current Payroll Month (auto)', '=IF(YEAR(TODAY())>C8,"December",IF(YEAR(TODAY())<C8,"January",CHOOSE(MONTH(TODAY()),' + MONTHS.map(m => '"' + m + '"').join(',') + ')))'],
    ['Company Logo URL (optional, public image link)', ''],
    ['Default EPF – Employer Share (flat ₹/month)', 1800],
    ['Employer EPF / ESI share borne from employee CTC?', 'Yes'],
    ['Sandwich Leave Rule', 'Yes'],
    ['Leave days bordering a sandwich are', 'Unpaid (LOP)']
  ];
  comp.forEach((row, i) => {
    const r = 4 + i;
    sh.getRange(r, 1, 1, 2).merge().setValue(row[0]).setFontWeight('bold').setFontColor(APP.TEXT).setBackground(APP.PALE);
    const v = sh.getRange(r, 3, 1, 2).merge();
    if (String(row[1]).charAt(0) === '=') v.setFormula(row[1]); else v.setValue(row[1]);
    v.setBackground(APP.INPUT).setHorizontalAlignment('left');
  });
  sh.getRange('C14:D14').setBackground(APP.AUTO);
  sh.getRange('C13:D13').setNumberFormat('dd-mmm-yyyy');
  sh.getRange('C8:D8').setNumberFormat('0').setFontWeight('bold').setFontSize(12).setFontColor(APP.RED);
  box_(sh.getRange('A4:D19'));
  sh.getRange('C18').setNote('Yes = weekly offs / holidays that fall between two Leave/Absent days (e.g. Absent Friday + Absent Monday) are also deducted. Which statuses trigger it is set in the Attendance Status table (column N).');
  sh.getRange('C19').setNote('Unpaid (LOP) = a paid status such as Leave on the Friday/Monday of a sandwich also becomes unpaid, so Leave Fri + Sat + Sun + Leave Mon = 4 days deducted.\nPaid as per status % = only the days in between are deducted.');
  sh.getRange('C8').setNote('Change this ONE cell (e.g. 2027) and every date, weekday, February length and calculation updates automatically.');
  sh.getRange('C13').setNote('MISSING attendance is only flagged on or after this date. Set it to the day you start using the system.');
  sh.getRange('C11').setNote(M30 + ': Daily salary = Monthly ÷ 30 (weekly offs & holidays are paid).\n' + MWD + ': Daily salary = Monthly ÷ working days of the employee\'s shift in that month.');
  sh.getRange('C12').setNote('Employee share of EPF, used when an employee has EPF Applicable = Yes but no amount of their own. Flat rupees, not a %.');
  sh.getRange('C16').setNote('Employer share of EPF, used when an employee has no employer-share amount of their own.');
  sh.getRange('C17').setNote('Yes = the Monthly Salary is treated as CTC: the employer EPF share is part of it and is shown on the slip as a deduction (total EPF deducted = employee + employer share, e.g. ₹1,800 + ₹1,800 = ₹3,600).\nNo = the company pays its share on top; only the employee share is deducted.');

  // ---------- B. Shift settings ----------
  section_(sh.getRange('F3:J3'), 'B.  SHIFT SETTINGS');
  header_(sh.getRange('F4:J4'), ['Shift Name', 'Start Time', 'End Time', 'Crosses Midnight?', 'Shift Hours']);
  sh.getRange('F5:H8').setValues([['Day', 9 / 24, 17 / 24], ['Night', 21 / 24, 6.5 / 24], ['Day (5-day)', 9 / 24, 17 / 24], ['', '', '']]);
  for (let r = 5; r <= 8; r++) {
    sh.getRange('I' + r).setFormula(`=IF(F${r}="","",IF(H${r}<G${r},"Yes","No"))`);
    sh.getRange('J' + r).setFormula(`=IF(F${r}="","",ROUND(MOD(H${r}-G${r},1)*24,2))`);
  }
  sh.getRange('F5:H8').setBackground(APP.INPUT);
  sh.getRange('I5:J8').setBackground(APP.AUTO).setHorizontalAlignment('center');
  sh.getRange('G5:H8').setNumberFormat('h:mm AM/PM').setHorizontalAlignment('center');
  sh.getRange('F7').setNote('For day-shift staff who get BOTH Saturday and Sunday off. Assign this shift to them in the Employee Master (column E).');
  box_(sh.getRange('F4:J8'));

  // ---------- D. Weekly offs (per shift) ----------
  section_(sh.getRange('F10:J10'), 'D.  WEEKLY OFF DAYS  (per shift)');
  sh.getRange('F11').setValue('Day of Week');
  for (let k = 0; k < 4; k++) sh.getRange(11, 7 + k).setFormula(`=IF($F$${5 + k}="","(no shift)",$F$${5 + k}&" shift")`);
  styleHeader_(sh.getRange('F11:J11'));
  const offs = DAYNAMES.map(d => [d,
    d === 'Sunday' ? 'Yes' : 'No',                                   // Day shift (6-day week)
    (d === 'Saturday' || d === 'Sunday') ? 'Yes' : 'No',             // Night shift
    (d === 'Saturday' || d === 'Sunday') ? 'Yes' : 'No',             // Day (5-day)
    (d === 'Saturday' || d === 'Sunday') ? 'Yes' : 'No']);
  sh.getRange('F12:J18').setValues(offs);
  sh.getRange('F12:F18').setFontWeight('bold').setBackground(APP.PALE);
  sh.getRange('G12:J18').setBackground(APP.INPUT).setHorizontalAlignment('center')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).setAllowInvalid(false).build());
  box_(sh.getRange('F11:J18'));

  // ---------- C. Attendance status ----------
  section_(sh.getRange('L3:N3'), 'C.  ATTENDANCE STATUS & PAY %');
  header_(sh.getRange('L4:N4'), ['Attendance Status', 'Paid Percentage', 'Triggers Sandwich?']);
  sh.getRange('L5:N9').setValues([
    ['Present', 1, 'No'],
    ['WFH', 0.5, 'No'],
    ['Half Day', 0.5, 'No'],
    ['Leave', 1, 'Yes'],
    ['Absent', 0, 'Yes']]);
  sh.getRange('L5:L9').setFontWeight('bold');
  sh.getRange('M5:M9').setNumberFormat('0%').setHorizontalAlignment('center').setBackground(APP.INPUT)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(0, 1).setAllowInvalid(false)
      .setHelpText('Enter a % between 0% and 100%').build());
  sh.getRange('N5:N9').setHorizontalAlignment('center').setBackground(APP.INPUT)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).setAllowInvalid(false).build());
  sh.getRange('N4').setNote('Yes = this status on BOTH sides of a weekly off / holiday makes the days in between unpaid (sandwich leave).');
  for (let i = 0; i < 5; i++) sh.getRange(5 + i, 12).setBackground(STATUS_COLORS[i][0]).setFontColor(STATUS_COLORS[i][1]);
  box_(sh.getRange('L4:N9'));

  // ---------- Rules note ----------
  sh.getRange('L11:N19').merge().setValue(
    'HOW THE RULES WORK\n' +
    '• Night shift (e.g. 9:00 PM → 6:30 AM) is ONE attendance entry, marked on the date the shift STARTS. The next morning is never a separate day.\n' +
    '• Weekly offs, holidays, dates before joining and future dates are never Absent/Missing and never deducted.\n' +
    '• Paid % drives every salary figure — WFH and Half Day are 50% by default; change them and all months, the Dashboard and slips update.\n' +
    '• EPF is flat: Applicable = No → ₹0; Yes → employee (and, if C17 = Yes, employer) share from the Employee Master, else the defaults.\n' +
    '• Weekly offs & holidays fill in automatically as "Weekly Off" / "Holiday". Lock finished months: Payroll System ▸ Finalize.\n' +
    '• SANDWICH: Leave/Absent on both sides of a weekly off or holiday (e.g. Fri + Mon) → the days in between are deducted too (C18/C19).\n' +
    '• MISSING = a past working day with no status (from the Tracking Start Date).')
    .setWrap(true).setVerticalAlignment('top').setBackground(APP.RED_LIGHT).setFontColor(APP.TEXT).setFontSize(9);
  box_(sh.getRange('L11:N19'), APP.RED);

  // ---------- E. Holiday list ----------
  section_(sh.getRange('AB3:AD3'), 'E.  COMPANY HOLIDAY CALENDAR  (enter at the start of the year)');
  header_(sh.getRange('AB4:AD4'), ['Holiday Date', 'Holiday Name', 'Applies To']);
  const hol = sh.getRange(5, APP.HOL_COL, APP.HOL_ROWS, 3);
  hol.setBackground(APP.INPUT);
  sh.getRange(5, APP.HOL_COL, APP.HOL_ROWS, 1).setNumberFormat('dd-mmm-yyyy (ddd)')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText('Enter a valid date').build());
  sh.getRange(5, APP.HOL_COL + 2, APP.HOL_ROWS, 1).setHorizontalAlignment('center')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(sh.getRange('AI12:AI16'), true).setAllowInvalid(false)
      .setHelpText('All shifts, or one shift only (blank = all shifts)').build());
  sh.getRange('AD4').setNote('All = every shift is off. Choose a shift name if only that shift gets the holiday. Blank = All.');
  box_(sh.getRange(4, APP.HOL_COL, APP.HOL_ROWS + 1, 3));

  // ---------- Active list helper ----------
  section_(sh.getRange('AF3'), 'ACTIVE EMPLOYEES (auto)');
  sh.getRange('AF4').setValue('Feeds the Salary Slip dropdown').setFontColor(APP.MUTED).setFontStyle('italic');
  sh.getRange('AF5').setFormula('=IFERROR(FILTER(Emp_ID&" - "&Emp_Name,Emp_ID<>"",Emp_Status<>"Inactive"),"")');
  sh.getRange('AF5:AF' + APP.MASTER_LAST).setBackground(APP.AUTO).setFontColor(APP.MUTED);
  // attendance dropdown list = 5 statuses + Weekly Off + Holiday
  section_(sh.getRange('AH3:AI3'), 'AUTO LISTS');
  sh.getRange('AH4:AI4').setValues([['Day-cell dropdown', 'Sandwich statuses']]).setFontColor(APP.MUTED).setFontSize(8);
  sh.getRange('AI5:AI9').setFormulas([5, 6, 7, 8, 9].map(r => [`=IF(N${r}="Yes",L${r},"#none#")`])).setBackground(APP.AUTO).setFontColor(APP.MUTED);
  sh.getRange('AI11').setValue('Holiday applies-to').setFontColor(APP.MUTED).setFontSize(8);
  sh.getRange('AI12:AI16').setFormulas([['="All"'], ['=F5&""'], ['=F6&""'], ['=F7&""'], ['=F8&""']]).setBackground(APP.AUTO).setFontColor(APP.MUTED);
  sh.getRange('AH5:AH11').setFormulas([['=L5'], ['=L6'], ['=L7'], ['=L8'], ['=L9'], ['="' + OFF_LABEL + '"'], ['="' + HOL_LABEL + '"']])
    .setBackground(APP.AUTO).setFontColor(APP.MUTED);

  // ---------- F. Employee Master ----------
  const MF = APP.MASTER_FIRST, ML = APP.MASTER_LAST;
  section_(sh.getRange('A21:Z21'), 'F.  EMPLOYEE MASTER  (salary structure in columns R–Y)');
  sh.getRange('A22:Z22').merge().setValue('SAMPLE DATA — Replace with Actual Employees   (or use menu: Payroll System ▸ Clear Sample Data)')
    .setBackground(APP.RED).setFontColor(APP.WHITE).setFontWeight('bold').setHorizontalAlignment('center');
  header_(sh.getRange('A23:Z23'), ['Employee ID', 'Employee Name', 'Department', 'Designation', 'Shift (base)', 'Joining Date',
    'Monthly Gross Salary (auto = sum of heads)', 'Salary Calculation Method', "Daily Salary (today's rate, auto)", 'EPF Applicable', 'EPF – Employee Share',
    'EPF – Employer Share', 'Other Deduction (monthly)', 'Email', 'Employment Status', 'Current Shift (auto)', 'Current Monthly Salary (auto)',
    'Basic', 'Dearness Allowance', 'Special Allowance', 'Laptop Allowance', 'CEA', 'Others', 'Medical Allowance',
    'ESI Employer (flat)', 'Est. Monthly In-hand (auto)']);
  sh.getRange('R23:X23').setBackground(APP.GREEN);   // salary heads — header text can be renamed; the slip uses these names
  sh.getRange('R23').setNote('Salary heads: rename any header (R23:X23) and the Salary Slip uses the new name.');
  sh.setRowHeight(23, 42);
  sh.getRange('A23:Z23').setWrap(true);
  sh.getRange(MF, 1, APP.ROWS, 15).setBackground(APP.INPUT);
  sh.getRange(MF, 18, APP.ROWS, 8).setBackground(APP.INPUT);
  sh.getRange(MF, 7, APP.ROWS, 1).setBackground(APP.AUTO).setFontColor(APP.NAVY).setFontWeight('bold');
  sh.getRange(MF, 26, APP.ROWS, 1).setBackground(APP.AUTO).setFontColor(APP.GREEN).setFontWeight('bold');
  sh.getRange(MF, 16, APP.ROWS, 2).setBackground(APP.AUTO).setFontColor(APP.NAVY).setFontWeight('bold');
  sh.getRange(MF, 9, APP.ROWS, 1).setBackground(APP.AUTO).setFontColor(APP.MUTED);
  const iF = [], pqF = [];
  for (let r = MF; r <= ML; r++) {
    const meth = `IF(OR(H${r}="",H${r}="${MDEF}"),SalMethod,H${r})`;
    const srow = `(8+IFERROR(MATCH(P${r},ShiftNames,0),1))`;
    iF.push([`=IF(A${r}="","",IFERROR(IF(${meth}="${M30}",Q${r}/30,Q${r}/COUNTIF(INDIRECT("'"&CurMonthName&"'!H"&${srow}&":AL"&${srow}),"Work")),0))`]);
    const mx = `MAXIFS(ChgDate,ChgID,A${r},ChgDate,"<="&TODAY())`;
    pqF.push([`=IF(A${r}="","",IFERROR(IF(${mx}=0,E${r},INDEX(ChgShift,MATCH(A${r}&"|"&${mx},ChgKey,0))),E${r}))`,
      `=IF(A${r}="","",N(G${r})+SUMIFS(IncAmt,IncID,A${r},IncDate,"<="&TODAY()))`]);
  }
  sh.getRange(MF, 9, APP.ROWS, 1).setFormulas(iF);
  const gF = [], zF = [];
  for (let r = MF; r <= ML; r++) {
    gF.push([`=IF(A${r}="","",SUM(R${r}:X${r}))`]);
    zF.push([`=IF(A${r}="","",Q${r}-IF(J${r}="Yes",IF(N(K${r})>0,K${r},N(DefaultEPS))+IF(ErFromEmp="Yes",IF(N(L${r})>0,L${r},N(DefaultEPSEr)),0),0)-IF(ErFromEmp="Yes",N(Y${r}),0)-N(M${r}))`]);
  }
  sh.getRange(MF, 7, APP.ROWS, 1).setFormulas(gF);
  sh.getRange(MF, 26, APP.ROWS, 1).setFormulas(zF);
  sh.getRange(MF, 16, APP.ROWS, 2).setFormulas(pqF);
  sh.getRange(MF, 6, APP.ROWS, 1).setNumberFormat('dd-mmm-yyyy')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  sh.getRange(MF, 5, APP.ROWS, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(sh.getRange('F5:F8'), true).setAllowInvalid(false).build());
  sh.getRange(MF, 8, APP.ROWS, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList([MDEF, M30, MWD], true).setAllowInvalid(false).build());
  sh.getRange(MF, 10, APP.ROWS, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'], true).setAllowInvalid(false).build());
  sh.getRange(MF, 15, APP.ROWS, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive'], true).setAllowInvalid(false).build());
  const numRule = SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).setHelpText('Enter an amount (0 or more)').build();
  [11, 12, 13, 18, 19, 20, 21, 22, 23, 24, 25].forEach(c => sh.getRange(MF, c, APP.ROWS, 1).setDataValidation(numRule));
  sh.getRange(MF, 14, APP.ROWS, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireTextIsEmail().setAllowInvalid(true).build());
  sh.getRange(MF, 5, APP.ROWS, 1).setHorizontalAlignment('center');
  sh.getRange(MF, 10, APP.ROWS, 1).setHorizontalAlignment('center');
  sh.getRange(MF, 15, APP.ROWS, 1).setHorizontalAlignment('center');
  box_(sh.getRange(23, 1, APP.ROWS + 1, 26));
  sh.getRange('E23').setNote('Shift the employee started on. Later moves go in section G (Shift Changes) with an effective date — do not overwrite this.');
  sh.getRange('G23').setNote('= Basic + DA + all allowances (columns R–X). Salary before any increment; increments go in section H with an effective date.');
  sh.getRange('Y23').setNote('Employer ESI contribution, flat ₹/month. Deducted from salary only when Settings C17 = Yes (CTC basis). ESI applies only when gross wages are ≤ ₹21,000.');
  sh.getRange('Z23').setNote('Current salary − EPF (employee + employer if CTC) − ESI (if CTC) − other deduction, for a full month with no LOP.');
  sh.getRange('P23').setNote('Shift in force today = base shift + latest Shift Change on or before today.');
  sh.getRange('Q23').setNote('Salary in force today = base salary + all increments effective on or before today.');
  sh.getRange('H23').setNote('"Company Default" follows the Salary Calculation Method in Company Settings. Choose a specific method to override it for this employee only.');
  sh.getRange('K23').setNote('Employee EPF share — flat rupees (not %). Blank = Default EPF – Employee Share. Ignored when EPF Applicable = No.');
  sh.getRange('L23').setNote('Employer EPF share — flat rupees. Blank = Default EPF – Employer Share. Deducted from salary only when Settings C17 = Yes (CTC basis).');
  sh.getRange('M23').setNote('Fixed monthly deduction (e.g. advance recovery). Applied in months with attendance.');
  sh.getRange('O23').setNote('Only Active employees are added to month sheets and the Salary Slip dropdown.');
  // CF: inactive rows grey, duplicate IDs red
  const mr = sh.getRange(MF, 1, APP.ROWS, 26);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=COUNTIF($A$${MF}:$A$${ML},$A${MF})>1`)
      .setBackground('#F28B82').setRanges([sh.getRange(MF, 1, APP.ROWS, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=$O${MF}="Inactive"`)
      .setBackground('#E0E0E0').setFontColor('#757575').setRanges([mr]).build()
  ]);
  sh.setFrozenRows(0);

  // ---------- G. Shift changes  &  H. Salary increments ----------
  const CF1 = APP.CHG_FIRST, CL1 = APP.CHG_LAST, CN = APP.CHG_ROWS;
  const GT = APP.CHG_TITLE;
  section_(sh.getRange(GT, 1, 1, 5), 'G.  SHIFT CHANGES  (with effective date)');
  sh.getRange(GT + 1, 1, 1, 5).merge().setValue('One row per move. From the effective date the new shift (and its weekly offs) applies; before it, the previous shift. Mid-month moves are handled day by day.')
    .setFontSize(8).setFontColor(APP.MUTED).setWrap(true);
  header_(sh.getRange(GT + 2, 1, 1, 5), ['Employee ID', 'Effective Date', 'New Shift', 'Remarks', 'Key (auto)']);
  section_(sh.getRange(GT, 8, 1, 5), 'H.  SALARY INCREMENTS  (flat amount, effective date)');
  sh.getRange(GT + 1, 8, 1, 5).merge().setValue('One row per increment. From the effective date salary = previous salary + increment. Mid-month increments are pro-rated day by day (use a minus amount for a reduction).')
    .setFontSize(8).setFontColor(APP.MUTED).setWrap(true);
  header_(sh.getRange(GT + 2, 8, 1, 5), ['Employee ID', 'Effective Date', 'Increment (flat ₹)', 'Remarks', 'New Monthly Salary (auto)']);
  sh.setRowHeight(GT + 1, 30);
  sh.getRange(CF1, 1, CN, 4).setBackground(APP.INPUT);
  sh.getRange(CF1, 8, CN, 4).setBackground(APP.INPUT);
  sh.getRange(CF1, 5, CN, 1).setBackground(APP.AUTO).setFontColor(APP.MUTED);
  sh.getRange(CF1, 12, CN, 1).setBackground(APP.AUTO).setFontColor(APP.NAVY).setFontWeight('bold');
  const keyF = [], newSalF = [];
  for (let r = CF1; r <= CL1; r++) {
    keyF.push([`=IF(OR(A${r}="",B${r}=""),"",A${r}&"|"&INT(B${r}))`]);
    newSalF.push([`=IF(OR(H${r}="",I${r}=""),"",IFERROR(INDEX(Emp_Salary,MATCH(H${r},Emp_ID,0))+SUMIFS(IncAmt,IncID,H${r},IncDate,"<="&I${r}),""))`]);
  }
  sh.getRange(CF1, 5, CN, 1).setFormulas(keyF);
  sh.getRange(CF1, 12, CN, 1).setFormulas(newSalF);
  const idRule = SpreadsheetApp.newDataValidation().requireValueInRange(sh.getRange(MF, 1, APP.ROWS, 1), true).setAllowInvalid(false).setHelpText('Pick an Employee ID from the Employee Master').build();
  const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText('Enter the effective date').build();
  sh.getRange(CF1, 1, CN, 1).setDataValidation(idRule);
  sh.getRange(CF1, 8, CN, 1).setDataValidation(idRule);
  sh.getRange(CF1, 2, CN, 1).setDataValidation(dateRule).setNumberFormat('dd-mmm-yyyy');
  sh.getRange(CF1, 9, CN, 1).setDataValidation(dateRule).setNumberFormat('dd-mmm-yyyy');
  sh.getRange(CF1, 3, CN, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(sh.getRange('F5:F8'), true).setAllowInvalid(false).build());
  sh.getRange(CF1, 10, CN, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(-10000000, 10000000).setAllowInvalid(false).setHelpText('Flat monthly increment in rupees (negative = reduction)').build());
  box_(sh.getRange(GT + 2, 1, CN + 1, 5)); box_(sh.getRange(GT + 2, 8, CN + 1, 5));

  // ---------- Named ranges ----------
  const ss = sh.getParent();
  defineNamedRanges_(ss);
  masterV2_(sh);                 // v2: Terminated / Discontinued + Last Working Day
  buildAdvanceSection_(sh);      // v2: I. Advance & Loan Register

  // company settings validations
  sh.getRange('C8').setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(2000, 2100).setAllowInvalid(false).setHelpText('Enter a year, e.g. 2026').build());
  sh.getRange('C11').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([M30, MWD], true).setAllowInvalid(false).build());
  sh.getRange('C12').setDataValidation(numRule);
  sh.getRange('C16').setDataValidation(numRule);
  sh.getRange('C17').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).setAllowInvalid(false).build());
  sh.getRange('C18').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).setAllowInvalid(false).build());
  sh.getRange('C19').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Unpaid (LOP)', 'Paid as per status %'], true).setAllowInvalid(false).build());
  sh.getRange('C13').setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
}

/** Creates (or re-creates) every named range used by the formulas. Safe to run any time. */
function defineNamedRanges_(ss) {
  const sh = ss.getSheetByName('Settings');
  if (!sh) return false;
  const MF = APP.MASTER_FIRST, ML = APP.MASTER_LAST, CF1 = APP.CHG_FIRST, CL1 = APP.CHG_LAST;
  const nr = (n, a1) => ss.setNamedRange(n, sh.getRange(a1));
  nr('CompanyName', 'C4'); nr('CompanyAddress', 'C5'); nr('CompanyPhone', 'C6'); nr('CompanyEmail', 'C7');
  nr('AttYear', 'C8'); nr('CurrencySym', 'C9'); nr('WorkHours', 'C10'); nr('SalMethod', 'C11');
  nr('DefaultEPS', 'C12'); nr('TrackStart', 'C13'); nr('CurMonthName', 'C14'); nr('LogoURL', 'C15');
  nr('ShiftNames', 'F5:F8'); nr('ShiftStart', 'G5:G8'); nr('ShiftEnd', 'H5:H8'); nr('ShiftCross', 'I5:I8');
  nr('WeeklyOffMatrix', 'G12:J18');
  nr('StatusNames', 'L5:L9'); nr('StatusPct', 'M5:M9');
  nr('HolidayDates', 'AB5:AB' + (4 + APP.HOL_ROWS)); nr('HolApplies', 'AD5:AD' + (4 + APP.HOL_ROWS));
  nr('SandwichOn', 'C18'); nr('SandwichBorder', 'C19'); nr('SandwichList', 'AI5:AI9'); nr('HeadNames', 'R23:X23');
  nr('ChgID', `A${CF1}:A${CL1}`); nr('ChgDate', `B${CF1}:B${CL1}`); nr('ChgShift', `C${CF1}:C${CL1}`); nr('ChgKey', `E${CF1}:E${CL1}`);
  nr('IncID', `H${CF1}:H${CL1}`); nr('IncDate', `I${CF1}:I${CL1}`); nr('IncAmt', `J${CF1}:J${CL1}`);
  nr('DefaultEPSEr', 'C16'); nr('ErFromEmp', 'C17'); nr('DropList', 'AH5:AH11');
  const cols = { Emp_ID: 'A', Emp_Name: 'B', Emp_Dept: 'C', Emp_Desig: 'D', Emp_Shift: 'E', Emp_Join: 'F',
    Emp_Salary: 'G', Emp_Method: 'H', Emp_Daily: 'I', Emp_EPSApp: 'J', Emp_EPSAmt: 'K', Emp_EPSErAmt: 'L',
    Emp_OtherDed: 'M', Emp_Email: 'N', Emp_Status: 'O', Emp_CurShift: 'P', Emp_CurSalary: 'Q',
    Emp_H1: 'R', Emp_H2: 'S', Emp_H3: 'T', Emp_H4: 'U', Emp_H5: 'V', Emp_H6: 'W', Emp_H7: 'X', Emp_ESI: 'Y', Emp_InHand: 'Z', Emp_Exit: 'AA' };
  Object.keys(cols).forEach(k => nr(k, `${cols[k]}${MF}:${cols[k]}${ML}`));
  nr('ActiveList', 'AF5:AF' + ML);
  const slip = ss.getSheetByName('Salary Slip'); if (slip) ss.setNamedRange('SlipPrintArea', slip.getRange('A4:G41'));
  return true;
}
/** Returns a named range; if it has gone missing, rebuilds the names from the Settings tab and tries again. */
function named_(ss, name) {
  let r = ss.getRangeByName(name);
  if (!r && defineNamedRanges_(ss)) r = ss.getRangeByName(name);
  if (!r) throw new Error('The workbook is not fully set up yet (the Settings tab is missing). Open Extensions ▸ Apps Script, run setupPayrollSystem once and wait for "Setup complete".');
  return r;
}

/* ============================== MONTH SHEETS =============================== */
function buildMonth_(sh, m) {
  const F = APP.M_FIRST, L = APP.M_LAST, N = APP.ROWS;
  sh.clear();
  ensureSize_(sh, L + 5, OUT_LAST + 3);
  sh.setTabColor(['#3B82F6', '#6366F1', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6', '#14B8A6', '#F97316', '#0EA5E9', '#84CC16', '#EC4899', '#64748B'][m - 1]);
  sh.setHiddenGridlines(true);

  // column widths
  const w = { A: 78, B: 150, C: 105, D: 120, E: 58, F: 92, G: 80 };
  Object.keys(w).forEach(k => sh.setColumnWidth(colN_(k), w[k]));
  sh.setColumnWidths(colN_('H'), 31, 42);
  sh.setColumnWidths(colN_('AM'), 6, 62);
  sh.setColumnWidths(colN_('AS'), 5, 80);
  sh.setColumnWidths(colN_('AX'), 8, 104);
  sh.setColumnWidth(colN_('BF'), 280);
  sh.setColumnWidths(colN_('BG'), SWB.LAST - colN_('BG') + 1, 70);

  // Row 1 title
  // (no merges across columns A|B → C, so the first two columns can be frozen)
  sh.getRange('A1').setFormula(`=CompanyName&"   |   ${m.toString().padStart(2, '0')} · ${MONTHS[m - 1].toUpperCase()} "&AttYear&"   —   ATTENDANCE & PAYROLL REGISTER"`);
  styleBanner_(sh.getRange('A1:AF1'), 14);
  sh.setRowHeight(1, 36);

  // Row 2 info
  sh.getRange('A2:G2').setValues([['Month No.', m, 'Year', '', 'Days in Month', '', 'Working days']]);
  sh.getRange('D2').setFormula('=AttYear');
  sh.getRange('F2').setFormula('=DAY(EOMONTH(DATE($D$2,$B$2,1),0))');
  sh.getRange('H2:AF2').merge().setFormula(
    '=TEXTJOIN("     |     ",TRUE,' + [9, 10, 11, 12].map(r => `IF($G$${r}="","",$G$${r}&" shift: "&COUNTIF($H$${r}:$AL$${r},"Work"))`).join(',') +
    ')&"          •  Default salary method: "&SalMethod');
  sh.getRange('A2:AF2').setFontColor(APP.MUTED).setFontSize(9);
  sh.getRange('B2').setFontWeight('bold'); sh.getRange('D2').setFontWeight('bold'); sh.getRange('F2').setFontWeight('bold');

  // Rows 3-4 monthly summary
  const S = (a, b) => `${a}${F}:${b || a}${L}`;
  const summary = [
    ['A3:B3', 'Active Employees', 'A4:B4', `=COUNTIFS(${S('A')},"<>",${S('BI')},"Active")`, '0'],
    ['C3', '=INDEX(StatusNames,1)', 'C4', `=SUM(${S('AM')})`, '0.##'],
    ['D3', '=INDEX(StatusNames,2)', 'D4', `=SUM(${S('AN')})`, '0.##'],
    ['E3', '=INDEX(StatusNames,3)', 'E4', `=SUM(${S('AO')})`, '0.##'],
    ['F3', '=INDEX(StatusNames,4)', 'F4', `=SUM(${S('AP')})`, '0.##'],
    ['G3', '=INDEX(StatusNames,5)', 'G4', `=SUM(${S('AQ')})`, '0.##'],
    ['H3:J3', 'Attendance %', 'H4:J4', `=IFERROR(SUM(${S('AM')})/SUM(${S('BL')}),"")`, '0.0%'],
    ['K3:M3', 'Paid Attendance %', 'K4:M4', `=IFERROR(SUM(${S('BM')})/SUM(${S('BL')}),"")`, '0.0%'],
    ['N3:P3', 'Missing Entries', 'N4:P4', `=SUM(${S('AR')})`, '0'],
    ['Q3:S3', 'Gross Payroll', 'Q4:S4', `=SUM(${S('AZ')})`, 'money'],
    ['T3:V3', 'Absence Deduction', 'T4:V4', `=SUM(${S('BA')})`, 'money'],
    ['W3:Y3', 'EPF Total', 'W4:Y4', `=SUM(${S('BB')})`, 'money'],
    ['Z3:AB3', 'Other Deductions', 'Z4:AB4', `=SUM(${S('BC')})`, 'money'],
    ['AC3:AF3', 'NET PAYROLL', 'AC4:AF4', `=SUM(${S('BE')})`, 'money']
  ];
  summary.forEach(s => {
    const lab = sh.getRange(s[0]); if (s[0].indexOf(':') > 0) lab.merge();
    if (String(s[1]).charAt(0) === '=') lab.setFormula(s[1]); else lab.setValue(s[1]);
    lab.setBackground(APP.NAVY2).setFontColor(APP.WHITE).setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
    const val = sh.getRange(s[2]); if (s[2].indexOf(':') > 0) val.merge();
    val.setFormula(s[3]).setBackground(APP.WHITE).setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center').setFontColor(APP.NAVY);
    if (s[4] !== 'money') val.setNumberFormat(s[4]);
  });
  sh.getRange('AC4:AF4').setFontColor(APP.RED);
  sh.getRange('N4:P4').setFontColor(APP.RED);
  box_(sh.getRange('A3:AF4'));
  sh.setRowHeight(3, 30); sh.setRowHeight(4, 30);

  // Row 6 legend
  sh.getRange('A6').setValue('How to use: pick a status from the dropdown in each day cell.  Red = MISSING (past working day left blank).  Do not sort this sheet.')
    .setFontSize(8).setFontColor(APP.MUTED);
  sh.setRowHeight(6, 30);
  const legend = [['=INDEX(StatusNames,1)', STATUS_COLORS[0]], ['=INDEX(StatusNames,2)', STATUS_COLORS[1]], ['=INDEX(StatusNames,3)', STATUS_COLORS[2]],
  ['=INDEX(StatusNames,4)', STATUS_COLORS[3]], ['=INDEX(StatusNames,5)', STATUS_COLORS[4]], ['MISSING', [MISSING_BG, '#7F0000']],
  ['Weekly Off', [OFF_BG, '#374151']], ['Holiday', [HOL_BG, '#0F5132']], ['Sandwich (deducted)', [SANDWICH_BG, '#9A3412']], ['Before joining', [PREJOIN_BG, '#6B7280']], ['Not in month', [NOMONTH_BG, '#FFFFFF']]];
  legend.forEach((lg, i) => {
    const rg = sh.getRange(6, 8 + i * 3, 1, 3).merge();
    if (lg[0].charAt(0) === '=') rg.setFormula(lg[0]); else rg.setValue(lg[0]);
    rg.setBackground(lg[1][0]).setFontColor(lg[1][1]).setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  });

  // Rows 7-12 date header
  sh.getRange('A7:B8').merge().setValue('NIGHT SHIFT RULE: attendance is marked on the date the shift STARTS (e.g. 9 PM). The following morning is NOT a separate day.')
    .setWrap(true).setFontSize(8).setFontColor(APP.NAVY).setFontWeight('bold').setVerticalAlignment('middle').setBackground(APP.RED_LIGHT);
  sh.getRange('A9:B12').merge().setValue('Day type by shift ▸  Work = working day,  Off = weekly off,  Hol = holiday  (automatic from Settings)')
    .setWrap(true).setFontSize(8).setFontColor(APP.MUTED).setVerticalAlignment('middle');
  sh.getRange('C7:F12').setBackground(APP.PALE);
  sh.getRange('G7').setValue('Date ▸'); sh.getRange('G8').setValue('Day ▸');
  sh.getRange('G9:G12').setFormulas([[`=INDEX(ShiftNames,1)&""`], [`=INDEX(ShiftNames,2)&""`], [`=INDEX(ShiftNames,3)&""`], [`=INDEX(ShiftNames,4)&""`]]);
  sh.getRange('G7:G12').setFontWeight('bold').setHorizontalAlignment('right').setFontSize(8).setBackground(APP.PALE);

  const r7 = [], r8 = [], r9 = [[], [], [], []], r13 = [];
  for (let d = 1; d <= 31; d++) {
    const c = colL_(7 + d);
    r13.push(d);
    r7.push(`=IF(${c}$13<=$F$2,DATE($D$2,$B$2,${c}$13),"")`);
    r8.push(`=IF(${c}$7="","",TEXT(${c}$7,"dddd"))`);
    for (let k = 1; k <= 4; k++) {
      const hol = `COUNTIF(HolidayDates,${c}$7)-COUNTIFS(HolidayDates,${c}$7,HolApplies,"?*")+COUNTIFS(HolidayDates,${c}$7,HolApplies,"All")+COUNTIFS(HolidayDates,${c}$7,HolApplies,$G$${8 + k})`;
      r9[k - 1].push(`=IF(OR(${c}$7="",$G$${8 + k}=""),"",IF(${hol}>0,"Hol",IF(INDEX(WeeklyOffMatrix,WEEKDAY(${c}$7,2),${k})="Yes","Off","Work")))`);
    }
  }
  sh.getRange('H13:AL13').setValues([r13]);
  sh.getRange('H7:AL7').setFormulas([r7]).setNumberFormat('dd-mmm-yyyy').setTextRotation(90).setFontSize(8).setHorizontalAlignment('center');
  sh.getRange('H8:AL8').setFormulas([r8]).setTextRotation(90).setFontSize(8).setHorizontalAlignment('center');
  sh.getRange('H9:AL12').setFormulas(r9).setFontSize(8).setHorizontalAlignment('center').setFontColor(APP.MUTED);
  sh.setRowHeight(7, 80); sh.setRowHeight(8, 68);
  box_(sh.getRange('G7:AL12'));

  // Row 13 headers
  const hdrA = ['Employee ID', 'Employee Name', 'Department', 'Designation', 'Shift', 'Monthly Salary', 'Daily Salary'];
  sh.getRange('A13:G13').setValues([hdrA]);
  sh.getRange('AM13:AQ13').setFormulas([[1, 2, 3, 4, 5].map(k => `=INDEX(StatusNames,${k})`)]);
  sh.getRange('AR13:BU13').setValues([['MISSING', 'Working Days (eligible)', 'Paid Days', 'Unpaid Days', 'Attendance %', 'Paid Attendance %',
    'Earned Salary (for period)', 'Other Earnings (input)', 'Gross Salary Earned', 'Absence Deduction', 'EPF (total deducted)', 'Other Deduction',
    'Total Deductions', 'NET SALARY', 'Remarks',
    'h:Join Date', 'h:Salary Method', 'h:Emp Status', 'h:Eligible Cal Days', 'h:Eligible Work Days', 'h:Marked Days',
    'h:Paid-weighted Days', 'h:Shift Row', 'h:Shift Work Days', 'h:Today Expected', 'h:Today Entry',
    'h:EPF Employee Share', 'h:EPF Employer Share', 'h:Shift changes in month', 'h:Increments in month']]);
  const hh = []; for (let d = 1; d <= 31; d++) hh.push('h:Type ' + d); for (let d = 1; d <= 31; d++) hh.push('h:Rate ' + d);
  hh.push('h:ESI Employer', 'h:Sandwich Days', 'h:Sandwich ₹', 'h:Border Unpaid Days', 'h:Border ₹');
  for (let d = 1; d <= 31; d++) hh.push('h:Sandwich ' + d);
  sh.getRange(13, colN_('BV'), 1, hh.length).setValues([hh]);
  styleHeader_(sh.getRange(13, 1, 1, SWB.LAST));
  sh.getRange('H13:AL13').setFontSize(10);
  sh.getRange('BE13').setBackground(APP.RED);
  sh.getRange('AY13').setBackground('#B7791F');
  sh.setRowHeight(13, 46);
  sh.getRange('A13:BS13').setWrap(true);

  // Employee rows are filled by the calculation engine (recalcMonths_) — no heavy formulas here

  // formats
  sh.getRange(`A${F}:A${L}`).setFontWeight('bold').setFontColor(APP.NAVY);
  sh.getRange(`E${F}:E${L}`).setHorizontalAlignment('center');
  sh.getRange(`H${F}:AL${L}`).setHorizontalAlignment('center').setFontSize(8)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInRange(named_(sh.getParent(), 'DropList'), true).setAllowInvalid(false)
      .setHelpText('Choose: Present, WFH, Half Day, Leave or Absent  (Weekly Off / Holiday fill in automatically)').build());
  sh.getRange(`AM${F}:AU${L}`).setNumberFormat('0.##').setHorizontalAlignment('center');
  sh.getRange(`AR${F}:AR${L}`).setFontColor('#B3261E').setFontWeight('bold');
  sh.getRange(`AV${F}:AW${L}`).setNumberFormat('0.0%').setHorizontalAlignment('center');
  sh.getRange(`AY${F}:AY${L}`).setBackground(APP.INPUT)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).setHelpText('Optional: bonus / incentive / arrears for this month').build());
  sh.getRange(`BE${F}:BE${L}`).setFontWeight('bold').setFontColor(APP.NAVY);
  sh.getRange(`BF${F}:BF${L}`).setFontSize(8).setWrap(true);
  sh.getRange(`BG${F}:BG${L}`).setNumberFormat('dd-mmm-yyyy');
  sh.getRange(`A${F}:BF${L}`).setBorder(true, true, true, true, true, true, '#E2E6EE', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(`A13:BF${L}`).setBorder(true, true, true, true, null, null, APP.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(`AM13:AM${L}`).setBorder(null, true, null, null, null, null, APP.NAVY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(`H13:H${L}`).setBorder(null, true, null, null, null, null, APP.NAVY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeights(F, N, 22);

  sh.hideColumns(colN_('BG'), SWB.LAST - colN_('BG') + 1);   // helper columns BG..FO
  monthV2_(sh);
  sh.setFrozenRows(13);
  sh.setFrozenColumns(2);
}

/** Formulas for one employee row on a month sheet.
 *  Returns {BG:[B..G], AM:[AM..AX], AZ:[AZ..BS], BT:[BT..EE]} */
function monthRowFormulas_(r, m) {
  const A = `$A${r}`;
  const lk = nm => `IFERROR(INDEX(${nm},MATCH(${A},Emp_ID,0)),"")`;
  const MS = `DATE($D$2,$B$2,1)`, ME = `EOMONTH(DATE($D$2,$B$2,1),0)`;
  // shift in force on a date = base shift, or the latest Shift Change on/before that date
  const mx = d => `MAXIFS(ChgDate,ChgID,${A},ChgDate,"<="&${d})`;
  const shiftAt = d => `IFERROR(IF(${mx(d)}=0,${lk('Emp_Shift')},INDEX(ChgShift,MATCH(${A}&"|"&${mx(d)},ChgKey,0))),${lk('Emp_Shift')})`;
  const TB = `$BV${r}:$CZ${r}`;                       // day type for THIS employee on each day (Work/Off/Hol)
  const RB = `$DA${r}:$EE${r}`;                       // monthly salary rate in force on each day
  const D = `$H$7:$AL$7`;                             // dates of the month
  const RW = `$H${r}:$AL${r}`;                        // attendance entries
  const Z = f => `=IF(${A}="","",${f})`;
  const EL = `(${TB}="Work")*(${D}>=$BG${r})`;          // eligible working day mask
  const cnt = c => Z(`SUMPRODUCT((${RW}=${c}$13)*${EL})`);
  const NOST = `ISNA(MATCH(${RW},StatusNames,0))`;   // cell holds no attendance status (blank, Weekly Off, Holiday…)
  const missCond = `${NOST},${TB}="Work",${D}>=MAX($BG${r},TrackStart),${D}<TODAY()`;
  const is30 = `$BH${r}="${M30}"`;
  // unpaid money = Σ over marked working days of (1 − paid %) × that day's rate
  const unpaidAmt = [1, 2, 3, 4, 5].map(k => `(1-INDEX(StatusPct,${k}))*SUMPRODUCT((${RW}=INDEX(StatusNames,${k}))*${EL}*${RB})`).join('+');
  const BG = [
    Z(`IFERROR(INDEX(Emp_Name,MATCH(${A},Emp_ID,0)),"(removed from master)")`),
    Z(`${lk('Emp_Dept')}&""`),
    Z(`${lk('Emp_Desig')}&""`),
    Z(`${shiftAt(ME)}&""`),                                                                                           // E shift at month end
    Z(`N(${lk('Emp_Salary')})+SUMIFS(IncAmt,IncID,${A},IncDate,"<="&${ME})`),                                        // F salary at month end
    Z(`IFERROR(IF(${is30},$F${r}/30,$F${r}/$BO${r}),0)`)                                                              // G daily (month-end rate)
  ];
  const AM = [
    cnt('AM'), cnt('AN'), cnt('AO'), cnt('AP'), cnt('AQ'),
    Z(`IF($BI${r}<>"Active",0,SUMPRODUCT(${NOST}*(${TB}="Work")*(${D}>=MAX($BG${r},TrackStart))*(${D}<TODAY())))`),   // AR missing
    Z(`$BK${r}`),                                                                                                      // AS
    Z(`IF($BL${r}=0,0,IF(${is30},$BJ${r},$BK${r})-$AU${r})`),                                                          // AT paid days
    Z(`ROUND($BL${r}-$BM${r}+N($EG${r})+N($EI${r}),2)`),                                                                // AU unpaid (incl. sandwich)
    Z(`IF($BL${r}=0,"",$AM${r}/$BL${r})`),                                                                             // AV att %
    Z(`IF($BL${r}=0,"",$BM${r}/$BL${r})`),                                                                             // AW paid %
    // AX earned: day-wise pro-rata of the rate in force (full month at one rate = exactly the monthly salary)
    Z(`IF($BL${r}=0,0,ROUND(IF(${is30},SUMPRODUCT(ISNUMBER(${D})*(${D}>=$BG${r})*${RB})/$F$2,IFERROR(SUMPRODUCT(${EL}*${RB})/$BO${r},0)),0))`)
  ];
  const AZ = [
    Z(`$AX${r}+N($AY${r})`),                                                                                           // AZ gross
    Z(`IF($BL${r}=0,0,MIN($AX${r},ROUND(IFERROR((${unpaidAmt}+N($EH${r})+N($EJ${r}))/IF(${is30},30,$BO${r}),0),0)))`),  // BA absence + sandwich ded
    Z(`$BR${r}+$BS${r}`),                                                                                             // BB EPF total
    Z(`IF($BL${r}=0,0,N(${lk('Emp_OtherDed')}))`),                                                                     // BC other ded
    Z(`$BA${r}+$BB${r}+$BC${r}+N($EF${r})`),                                                                           // BD total ded (incl. ESI)
    Z(`$AZ${r}-$BD${r}`),                                                                                              // BE net
    Z(`IF($BJ${r}=0,"Not employed in this month (joins later)",IF(N($AR${r})>0,"MISSING ("&$AR${r}&"): "&TEXTJOIN(", ",TRUE,FILTER(TEXT(${D},"dd-mmm"),${missCond})),IF($BL${r}=0,IF($BI${r}="Active","Attendance not entered for this period","Inactive employee — no attendance"),IF($BI${r}="Active","OK","OK (employee now inactive)"))))&IF(N($EG${r})>0,"  • Sandwich: "&$EG${r}&" day(s) deducted","")&IF(N($BT${r})>0,"  • Shift changed this month","")&IF(N($BU${r})>0,"  • Salary revised this month","")`), // BF remarks
    Z(`IFERROR(N(INDEX(Emp_Join,MATCH(${A},Emp_ID,0))),0)`),                                                           // BG join
    Z(`IF(OR(${lk('Emp_Method')}="",${lk('Emp_Method')}="${MDEF}"),SalMethod,${lk('Emp_Method')})`),                   // BH method
    Z(`IFERROR(IF(INDEX(Emp_Status,MATCH(${A},Emp_ID,0))="","Active",INDEX(Emp_Status,MATCH(${A},Emp_ID,0))&""),"Inactive")`), // BI status
    Z(`SUMPRODUCT(ISNUMBER(${D})*(${D}>=$BG${r}))`),                                                                   // BJ eligible cal days
    Z(`SUMPRODUCT(${EL})`),                                                                                            // BK eligible work days
    Z(`SUM($AM${r}:$AQ${r})`),                                                                                         // BL marked
    Z(`$AM${r}*INDEX(StatusPct,1)+$AN${r}*INDEX(StatusPct,2)+$AO${r}*INDEX(StatusPct,3)+$AP${r}*INDEX(StatusPct,4)+$AQ${r}*INDEX(StatusPct,5)`), // BM
    Z(`IFERROR(MATCH(${shiftAt(MS)},$G$9:$G$12,0),1)`),                                                                 // BN shift row at month start
    Z(`COUNTIF(${TB},"Work")`),                                                                                        // BO shift work days
    Z(`IF(AND(YEAR(TODAY())=$D$2,MONTH(TODAY())=$B$2,$BI${r}="Active",$BG${r}<=TODAY()),IF(INDEX(${TB},1,DAY(TODAY()))="Work",1,0),0)`), // BP
    Z(`IF(AND(YEAR(TODAY())=$D$2,MONTH(TODAY())=$B$2),INDEX(${RW},1,DAY(TODAY()))&"","")`),                            // BQ
    Z(`IF($BL${r}=0,0,IF(${lk('Emp_EPSApp')}<>"Yes",0,IF(N(${lk('Emp_EPSAmt')})>0,N(${lk('Emp_EPSAmt')}),N(DefaultEPS))))`),        // BR EPF employee
    Z(`IF(OR($BL${r}=0,ErFromEmp<>"Yes"),0,IF(${lk('Emp_EPSApp')}<>"Yes",0,IF(N(${lk('Emp_EPSErAmt')})>0,N(${lk('Emp_EPSErAmt')}),N(DefaultEPSEr))))`) // BS EPF employer (CTC)
  ];
  const BT = [
    Z(`COUNTIFS(ChgID,${A},ChgDate,">"&${MS},ChgDate,"<="&${ME})`),                                                     // BT shift changes inside month
    Z(`COUNTIFS(IncID,${A},IncDate,">"&${MS},IncDate,"<="&${ME})`)                                                      // BU increments inside month
  ];
  for (let d = 1; d <= 31; d++) {                                                                                       // BV..CZ day type
    const c = colL_(7 + d);
    BT.push(Z(`IF(${c}$7="","",INDEX(${c}$9:${c}$12,IF($BT${r}=0,$BN${r},IFERROR(MATCH(${shiftAt(c + '$7')},$G$9:$G$12,0),1))))`));
  }
  for (let d = 1; d <= 31; d++) {                                                                                       // DA..EE day rate
    const c = colL_(7 + d);
    BT.push(Z(`IF(${c}$7="",0,IF($BU${r}=0,$F${r},N(${lk('Emp_Salary')})+SUMIFS(IncAmt,IncID,${A},IncDate,"<="&${c}$7)))`));
  }
  return { BG: BG, AM: AM, AZ: AZ, BT: BT };
}

/** Rebuilds attendance colours on all month sheets (also picks up renamed statuses). */
function refreshAttendanceFormatting() {
  const ss = SpreadsheetApp.getActive();
  const names = named_(ss, 'StatusNames').getValues().map(r => String(r[0]));
  const EXL = colL_(XC.EXIT);   // last working day (FR)
  const trackS = toSerial_(ss.getSheetByName('Settings').getRange('C13').getValue(), ss.getSpreadsheetTimeZone()) || 0;
  const F = APP.M_FIRST, L = APP.M_LAST;
  MONTHS.forEach(m => {
    const sh = ss.getSheetByName(m); if (!sh) return;
    const grid = sh.getRange(`H${F}:AL${L}`);
    const rules = [];
    names.forEach((n, i) => {
      if (!n) return;
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(n)
        .setBackground(STATUS_COLORS[i][0]).setFontColor(STATUS_COLORS[i][1]).setBold(true).setRanges([grid]).build());
    });
    const cf = (f, bg, fc) => { const b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(f).setBackground(bg).setRanges([grid]); if (fc) b.setFontColor(fc); rules.push(b.build()); };
    // v1.3: no INDIRECT in colour rules (INDIRECT on ~1,500 cells per sheet slowed every edit); tracking start date is embedded
    cf(`=AND(H${F}="",$A${F}<>"",INDEX($BV${F}:$CZ${F},1,COLUMN(H${F})-7)="Work",H$7<TODAY(),H$7>=MAX($BG${F},${trackS}),OR($BI${F}="Active",N($${EXL}${F})>0),OR(N($${EXL}${F})=0,H$7<=$${EXL}${F}))`, MISSING_BG, '#7F0000');
    cf(`=AND($A${F}<>"",INDEX($${colL_(SWB.FLAG1)}${F}:$${colL_(SWB.LAST)}${F},1,COLUMN(H${F})-7)=1)`, SANDWICH_BG, '#9A3412');
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(OFF_LABEL).setBackground(OFF_BG).setFontColor('#6B7280').setRanges([grid]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(HOL_LABEL).setBackground(HOL_BG).setFontColor('#0F5132').setRanges([grid]).build());
    cf(`=H$7=""`, NOMONTH_BG);
    cf(`=AND($A${F}<>"",H$7<$BG${F})`, PREJOIN_BG);
    cf(`=AND($A${F}<>"",N($${EXL}${F})>0,H$7>$${EXL}${F})`, PREJOIN_BG);   // after the last working day
    cf(`=AND($A${F}<>"",INDEX($BV${F}:$CZ${F},1,COLUMN(H${F})-7)="Hol")`, HOL_BG);
    cf(`=AND($A${F}<>"",INDEX($BV${F}:$CZ${F},1,COLUMN(H${F})-7)="Off")`, OFF_BG);
    // header rows
    const hdr = sh.getRange('H7:AL12');
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=H$7=""').setBackground(NOMONTH_BG).setRanges([hdr]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Hol').setBackground(HOL_BG).setFontColor('#0F5132').setRanges([sh.getRange('H9:AL12')]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Off').setBackground(OFF_BG).setRanges([sh.getRange('H9:AL12')]).build());
    // remarks
    const rem = sh.getRange(`BF${F}:BF${L}`);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('MISSING').setBackground('#FDE2E1').setFontColor('#B3261E').setBold(true).setRanges([rem]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('OK').setFontColor(APP.GREEN).setRanges([rem]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('Attendance not').setBackground('#FFF4E5').setFontColor('#8A5A00').setRanges([rem]).build());
    // missing count
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#FDE2E1').setRanges([sh.getRange(`AR${F}:AR${L}`)]).build());
    // inactive employee rows (names greyed)
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND($A${F}<>"",$BI${F}<>"Active")`).setFontColor('#9CA3AF').setRanges([sh.getRange(`A${F}:G${L}`)]).build());
    sh.setConditionalFormatRules(rules);
  });
}

/* ============================ EMPLOYEE SYNC ================================ */
/**
 * Adds every Active employee (joined on/before month end) to each month sheet.
 * Rows are NEVER re-ordered or deleted, so attendance always stays with the right person.
 * Rows of inactive / removed employees with no attendance in that month are hidden.
 */
function syncEmployees() {
  const ss = SpreadsheetApp.getActive();
  const st = ss.getSheetByName('Settings'); if (!st) return;
  const tz = ss.getSpreadsheetTimeZone();
  const year = Number(st.getRange('C8').getValue());
  const master = st.getRange(APP.MASTER_FIRST, 1, APP.ROWS, APP.EXIT_COL).getValues()
    .filter(r => String(r[0]).trim() !== '')
    .map(r => { const status = String(r[14]).trim() || 'Active';
      return { id: String(r[0]).trim(), join: toSerial_(r[5], tz), status: status,
        exit: EXITED.indexOf(status) >= 0 ? toSerial_(r[APP.EXIT_COL - 1], tz) : 0 }; });
  const byId = {}; master.forEach(e => byId[e.id] = e);
  // on payroll in a month = joined by month end AND (Active, or Terminated/Discontinued with last working day in/after the month start)
  const onPay = (e, ms, me) => (!e.join || e.join <= me) && (e.status === 'Active' || (EXITED.indexOf(e.status) >= 0 && e.exit > 0 && e.exit >= ms));
  const F = APP.M_FIRST, N = APP.ROWS;
  MONTHS.forEach((mName, i) => {
    const sh = ss.getSheetByName(mName); if (!sh || isLocked_(sh)) return;   // finalized months are never changed
    const ms = serial_(year, i + 1, 1), me = serial_(year, i + 2, 0);
    const ids = sh.getRange(F, 1, N, 1).getValues().map(r => String(r[0]).trim());
    const att = sh.getRange(F, 8, N, 31).getValues();
    const ay = sh.getRange(F, colN_('AY'), N, 1).getValues();
    const hasData = k => att[k].some(v => { const t = String(v).trim(); return t !== '' && t !== OFF_LABEL && t !== HOL_LABEL; }) || String(ay[k][0]).trim() !== '';
    let changed = false;
    // 1. remove employees who left BEFORE this month (Terminated / Discontinued) — only if nothing was entered for them
    ids.forEach((x, k) => {
      const e = byId[x];
      if (x && e && EXITED.indexOf(e.status) >= 0 && e.exit > 0 && e.exit < ms && !hasData(k)) { ids[k] = ''; changed = true; }
    });
    // 1b. remove IDs that are no longer in the Employee Master (deleted or renamed) — only if nothing was entered for them
    ids.forEach((x, k) => { if (x && !byId[x] && !hasData(k)) { ids[k] = ''; changed = true; } });
    // 2. add employees who are on payroll this month
    const present = {}; ids.forEach(x => { if (x) present[x] = true; });
    master.forEach(e => {
      if (present[e.id] || !onPay(e, ms, me)) return;
      const slot = ids.indexOf('');
      if (slot < 0) return;
      ids[slot] = e.id; present[e.id] = true; changed = true;
    });
    if (changed) sh.getRange(F, 1, N, 1).setValues(ids.map(x => [x]));
    sh.showRows(F, N);
    ids.forEach((x, k) => {
      if (!x) return;
      const e = byId[x];
      if ((!e || !onPay(e, ms, me)) && !hasData(k)) sh.hideRows(F + k);
    });
  });
}

/**
 * Writes "Weekly Off" / "Holiday" into the day cells of every employee according to their
 * shift's weekly offs and the Holiday List. Only blank cells (or old Off/Holiday marks) are
 * touched — attendance HR has entered is never overwritten. Runs automatically when the
 * Employee Master, weekly offs, holidays or the year change.
 */
function markOffDays() {
  const ss = SpreadsheetApp.getActive();
  const ctx = readContext_(ss);
  const F = APP.M_FIRST, N = APP.ROWS;
  const isMark = x => x === OFF_LABEL || x === HOL_LABEL;
  MONTHS.forEach((mName, mi) => {
    const sh = ss.getSheetByName(mName); if (!sh || isLocked_(sh)) return;
    const ids = sh.getRange(F, 1, N, 1).getValues();
    const rng = sh.getRange(F, 8, N, 31);
    const grid = rng.getValues();
    const dim = daysIn_(ctx.year, mi + 1);
    let changed = false;
    ids.forEach((r, i) => { try {
      const id = String(r[0]).trim(), e = ctx.master[id];
      for (let d = 1; d <= 31; d++) {
        const cur = String(grid[i][d - 1]).trim();
        let want = null;
        if (!id || !e || d > dim) { if (isMark(cur)) want = ''; }
        else {
          const s = serial_(ctx.year, mi + 1, d);
          const t = dayType_(ctx, s, shiftIdx_(ctx, shiftOn_(ctx, id, s)));
          if (s < e.join || (e.exit && s > e.exit)) { if (isMark(cur)) want = ''; }   // before joining / after exit
          else if (t === 'Off' && (cur === '' || cur === HOL_LABEL)) want = OFF_LABEL;
          else if (t === 'Hol' && (cur === '' || cur === OFF_LABEL)) want = HOL_LABEL;
          else if (t === 'Work' && isMark(cur)) want = '';
        }
        if (want !== null && want !== cur) { grid[i][d - 1] = want; changed = true; }
      }
    } catch (err) { /* v2.2: a bad row is skipped; payroll remarks show the problem */ } });
    if (changed) rng.setValues(grid);
  });
}

/* ============================ FINALIZE / UNLOCK ============================ */
/** Freezes a processed month: formulas become fixed values, so later changes to salary,
 *  EPF, shifts or settings can never alter payroll that has already been paid. */
function finalizeMonth() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Finalize (lock) a month', 'Type the month to lock, e.g. September.\nPayroll figures for that month will be frozen as values.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const m = monthFromText_(res.getResponseText()); if (!m) { ui.alert('Not a month name: ' + res.getResponseText()); return; }
  const sh = SpreadsheetApp.getActive().getSheetByName(m);
  if (isLocked_(sh)) { ui.alert(m + ' is already finalized.'); return; }
  if (ui.alert('Lock ' + m + '?', 'Salary, EPF and attendance figures of ' + m + ' will be frozen as they are now.\nYou can unlock later (Payroll System ▸ Unlock a month).', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  lockSheet_(sh);
  ui.alert('🔒 ' + m + ' finalized.');
}
function unlockMonth() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Unlock a month', 'Type the month to unlock, e.g. September.\nIts figures will be recalculated from CURRENT settings and employee data.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const m = monthFromText_(res.getResponseText()); if (!m) { ui.alert('Not a month name: ' + res.getResponseText()); return; }
  const sh = SpreadsheetApp.getActive().getSheetByName(m);
  if (!isLocked_(sh)) { ui.alert(m + ' is not locked.'); return; }
  unlockSheet_(sh);
  ui.alert('🔓 ' + m + ' unlocked and recalculated.');
}
function lockSheet_(sh) {
  SpreadsheetApp.flush();
  const F = APP.M_FIRST, L = APP.M_LAST;
  ['AR', 'BF', 'BP', 'BQ'].forEach(c => { const r = sh.getRange(`${c}${F}:${c}${L}`); r.setValues(r.getValues()); });   // freeze the few live formulas
  sh.getRange('AG1:AL1').merge().setValue('🔒 FINALIZED ' + Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd-MMM-yyyy'))
    .setBackground(APP.RED).setFontColor(APP.WHITE).setFontWeight('bold').setHorizontalAlignment('center');
}
function unlockSheet_(sh) {
  sh.getRange('AG1:AL1').breakApart().clearContent().setBackground(APP.NAVY);
  recalcMonths_([MONTHS.indexOf(sh.getName()) + 1]);
}
/* ========================= CALCULATION ENGINE ============================== */
/**
 * All payroll maths runs here in Apps Script and the results are written to the month sheets as
 * plain values (fast to open and edit). Only 4 light formulas per employee stay live because they
 * depend on today's date: MISSING count, Remarks, today's expected / entered status.
 * Runs automatically on every edit (installable trigger) — or Payroll System ▸ Recalculate Payroll Now.
 */
function recalcAll_() { recalcMonths_([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]); }
function recalculatePayrollNow() {
  recalcAll_(); renderRegister_();
  toast_('All 12 months recalculated ✔', 'Payroll', 5);
}

function recalcMonths_(months, onlyIds) {
  // FAST ENGINE: reads only the months it needs; with onlyIds it recalculates and writes ONLY those employees' rows.
  // v2: advance / loan recovery runs month by month — if a month's recovery changes, that employee's LATER months
  // are recalculated too (so balances always carry forward correctly).
  const lock = LockService.getDocumentLock();
  lock.waitLock(60000);
  try {
    const ss = SpreadsheetApp.getActive();
    const ctx = readContext_(ss);
    const F = APP.M_FIRST, N = APP.ROWS, AYC = colN_('AY');
    const cache = {};
    const getData = mi => {                                      // mi = 0-based month index, read once (lazy)
      if (!cache[mi]) {
        const sh = ss.getSheetByName(MONTHS[mi]);
        if (sh.getMaxColumns() < OUT_LAST + 1) ensureSize_(sh, Math.max(sh.getMaxRows(), APP.M_LAST + 5), OUT_LAST + 3);   // v2.2: self-repair missing columns
        const v = sh.getRange(F, 1, N, AYC).getValues();         // A..AY in ONE read
        const ids = v.map(r => String(r[0]).trim());
        const pos = {}; ids.forEach((id, i) => { if (id) pos[id] = i; });
        cache[mi] = { sh: sh, ids: ids, pos: pos, att: v.map(r => r.slice(7, 38).map(x => String(x).trim())), ay: v.map(r => r[AYC - 1]), locked: null };
      }
      return cache[mi];
    };
    const work = {};                                             // month (1-12) → 'ALL' or Set of employee IDs
    months.forEach(m => { work[m] = onlyIds ? new Set(onlyIds) : 'ALL'; getData(m - 1); });
    const typeCache = {};
    const typeOf = (id, s) => { const k = id + '|' + s; if (!(k in typeCache)) typeCache[k] = dayType_(ctx, s, shiftIdx_(ctx, shiftOn_(ctx, id, s))); return typeCache[k]; };
    const isSt = x => ctx.statuses.indexOf(x) >= 0;
    const lastWork = (id, mi) => { const D = getData(mi), i = D.pos[id]; if (i === undefined) return null;
      for (let d = daysIn_(ctx.year, mi + 1); d >= 1; d--) if (typeOf(id, serial_(ctx.year, mi + 1, d)) === 'Work') return { v: D.att[i][d - 1] }; return { none: true }; };
    const firstWork = (id, mi) => { const D = getData(mi), i = D.pos[id]; if (i === undefined) return null;
      for (let d = 1; d <= daysIn_(ctx.year, mi + 1); d++) if (typeOf(id, serial_(ctx.year, mi + 1, d)) === 'Work') return { v: D.att[i][d - 1] }; return { none: true }; };
    const prevStatus = (id, mi) => { for (let k = mi - 1; k >= 0; k--) { const x = lastWork(id, k); if (x === null) return ''; if (!x.none) return x.v; } return ''; };
    const nextStatus = (id, mi) => { for (let k = mi + 1; k <= 11; k++) { const x = firstWork(id, k); if (x === null) return ''; if (!x.none) return x.v; } return ''; };
    const W = OUT_LAST - OUT_FIRST + 1;
    const results = [];
    let advDirty = false;
    for (let m = 1; m <= 12; m++) {
      const w = work[m]; if (!w) continue;
      const D = getData(m - 1);
      if (D.locked === null) D.locked = isLocked_(D.sh);
      if (D.locked) continue;                                    // finalized months keep their figures (and recoveries)
      const one = (i, id) => {
        const known = !!ctx.master[id];
        let res;
        try { res = calcRow_(ctx, m, id, D.att[i], D.ay[i], typeOf, isSt, known ? prevStatus(id, m - 1) : '', known ? nextStatus(id, m - 1) : ''); }
        catch (err) {                                  // v2.2 safety net: one bad row never stops the rest of payroll
          const out = blankOut_(D.ay[i]);
          out[colN_('BF') - OUT_FIRST] = '⚠ Could not calculate this row: ' + (err && err.message ? err.message : err) + ' — check this employee in Settings';
          res = { bg: [(ctx.master[id] && ctx.master[id].name) || '(check employee)', '', '', '', 0, 0], am: out, advChanged: false };
        }
        if (res.advChanged) {
          advDirty = true;
          for (let k = m + 1; k <= 12; k++) if (work[k] !== 'ALL') { work[k] = work[k] || new Set(); work[k].add(id); }
        }
        return res;
      };
      if (w === 'ALL') {
        const bg = [], am = [];
        for (let i = 0; i < N; i++) {
          const id = D.ids[i];
          const res = id ? one(i, id) : null;
          bg.push(res ? res.bg : ['', '', '', '', '', '']);
          am.push(res ? res.am : blankOut_(D.ay[i]));
        }
        results.push({ sh: D.sh, row: F, n: N, bg: bg, am: am });
        // employees with an advance who are NOT on this month's sheet recover nothing this month
        Object.keys(ctx.adv).forEach(id => { if (D.pos[id] === undefined) ctx.adv[id].forEach(a => { if (a.rec[m - 1]) { a.rec[m - 1] = 0; advDirty = true; } }); });
      } else {
        w.forEach(id => {
          const i = D.pos[id]; if (i === undefined) return;
          const res = one(i, id);
          results.push({ sh: D.sh, row: F + i, n: 1, bg: [res.bg], am: [res.am] });
        });
      }
    }
    results.forEach(x => {                                       // write everything in one batch
      x.sh.getRange(x.row, 2, x.n, 6).setValues(x.bg);
      x.sh.getRange(x.row, OUT_FIRST, x.n, W).setValues(x.am);
    });
    if (ctx.hasAdv && (advDirty || !onlyIds)) writeAdvRegister_(ss, ctx);
  } finally { lock.releaseLock(); }
}

/** Writes the month-by-month recoveries (Jan..Dec) back into Settings ▸ I. Advance & Loan Register. */
function writeAdvRegister_(ss, ctx) {
  ss.getSheetByName('Settings').getRange(APP.ADV_FIRST, 14, APP.ADV_ROWS, 12)
    .setValues(ctx.advRows.map(r => r.map(x => (x === '' || !x) ? '' : round_(Number(x), 2))));
}

function blankOut_(ayVal) {
  const row = new Array(OUT_LAST - OUT_FIRST + 1).fill('');
  row[colN_('AY') - OUT_FIRST] = ayVal;
  return row;
}

/** Calculates one employee-month. Returns {bg:[B..G], am:[AM..FS], advChanged} — all plain values. */
function calcRow_(ctx, m, id, att, ayVal, typeOf, isSt, ps0, ns32) {
  const e = ctx.master[id];
  const out = new Array(OUT_LAST - OUT_FIRST + 1).fill('');
  const put = (col, v) => { out[colN_(col) - OUT_FIRST] = v; };
  const putN = (n, v) => { out[n - OUT_FIRST] = v; };
  const y = ctx.year, dim = daysIn_(y, m), ms = serial_(y, m, 1), me = serial_(y, m, dim);
  const tdy = ctx.today, inMonth = tdy >= ms && tdy <= me;
  put('AY', ayVal);
  put('BQ', inMonth ? (att[tdy - ms] || '') : '');
  put('BP', 0); put('AR', 0);
  if (!e) {                                              // ID no longer in the Employee Master
    put('BI', 'Inactive'); ['AM', 'AN', 'AO', 'AP', 'AQ', 'AS', 'AT', 'AU', 'AX', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BJ', 'BK', 'BL', 'BM', 'BO', 'BR', 'BS', 'EF', 'EG'].forEach(c => put(c, 0));
    put('AZ', N_(ayVal)); put('BE', N_(ayVal)); put('BN', 1);
    putN(XC.ADV, 0); putN(XC.NETPRE, N_(ayVal));
    put('BF', 'Removed from Employee Master — no payroll calculated');
    return { bg: ['(removed from master)', '', '', '', 0, 0], am: out, advChanged: false };
  }
  const exited = EXITED.indexOf(e.status) >= 0, exit = e.exit || 0;
  const join = e.join || 0;
  const elig = s => s >= join && (!exit || s <= exit);                                  // employed on that date
  const onPay = s => (e.status === 'Active' || (exited && exit > 0)) && elig(s);       // attendance expected on that date
  const method = (e.method === '' || e.method === MDEF) ? ctx.defMethod : e.method, is30 = method === M30;
  const F = rateOn_(ctx, id, me);
  const type = [], rate = [];
  for (let d = 1; d <= 31; d++) {
    if (d > dim) { type.push(''); rate.push(0); continue; }
    const s = ms + d - 1; type.push(typeOf(id, s)); rate.push(rateOn_(ctx, id, s));
  }
  const cnt = [0, 0, 0, 0, 0]; let eligCal = 0, eligWD = 0, shiftWD = 0, sumCal = 0, sumWD = 0, unpaidAmt = 0, missing = 0;
  const missDates = [];
  for (let d = 1; d <= dim; d++) {
    const s = ms + d - 1, t = type[d - 1], rt = rate[d - 1], en = att[d - 1];
    if (t === 'Work') shiftWD++;
    if (elig(s)) {
      eligCal++; sumCal += rt;
      if (t === 'Work') { eligWD++; sumWD += rt; const k = ctx.statuses.indexOf(en); if (k >= 0) { cnt[k]++; unpaidAmt += (1 - ctx.pct[k]) * rt; } }
    }
    if (onPay(s) && t === 'Work' && s >= ctx.trackStart && s < tdy && !isSt(en)) {
      missing++; missDates.push(String(d).padStart(2, '0') + '-' + MONTHS[m - 1].slice(0, 3));
    }
  }
  const marked = cnt[0] + cnt[1] + cnt[2] + cnt[3] + cnt[4];
  const paidW = cnt.reduce((a, c, k) => a + c * ctx.pct[k], 0);
  // ---- sandwich ----
  const trig = x => ctx.sandwich.indexOf(x) >= 0;
  const PS = [ps0], NS = new Array(33).fill('');
  for (let d = 1; d <= 31; d++) PS[d] = type[d - 1] === 'Work' ? att[d - 1] : PS[d - 1];
  NS[32] = ns32;
  for (let d = 31; d >= 1; d--) NS[d] = type[d - 1] === 'Work' ? att[d - 1] : NS[d + 1];
  const OFF = d => (d >= 1 && d <= 31 && (type[d - 1] === 'Off' || type[d - 1] === 'Hol')) ? 1 : 0;
  const swOn = ctx.sandwichOn === 'Yes';
  let swc = 0, swm = 0, brw = 0, brm = 0; const flags = [];
  for (let d = 1; d <= 31; d++) {
    const s = ms + d - 1;
    const sw = d <= dim && OFF(d) && !isSt(att[d - 1]) && trig(PS[d - 1]) && trig(NS[d + 1]) && elig(s);
    flags.push(sw && swOn ? 1 : 0);
    if (sw) { swc++; swm += rate[d - 1]; }
    if (d <= dim && type[d - 1] === 'Work' && elig(s) && trig(att[d - 1]) &&
        ((OFF(d + 1) && trig(NS[d + 1])) || (OFF(d - 1) && trig(PS[d - 1])))) {
      const pc = ctx.pct[ctx.statuses.indexOf(att[d - 1])]; brw += pc; brm += pc * rate[d - 1];
    }
  }
  if (!swOn || marked === 0) { swc = 0; swm = 0; }
  if (!swOn || marked === 0 || ctx.sandwichBorder !== 'Unpaid (LOP)') { brw = 0; brm = 0; }
  // ---- money ----
  const daily = is30 ? F / 30 : (shiftWD ? F / shiftWD : 0);
  const unpaid = round_(marked - paidW + swc + brw, 2);
  const earned = marked === 0 ? 0 : round_(is30 ? sumCal / dim : (shiftWD ? sumWD / shiftWD : 0), 0);
  const lopDiv = is30 ? 30 : shiftWD;
  const absence = marked === 0 ? 0 : Math.min(earned, lopDiv ? round_((unpaidAmt + swm + brm) / lopDiv, 0) : 0);
  const epfEE = (marked === 0 || e.epsApp !== 'Yes') ? 0 : (e.epsAmt > 0 ? e.epsAmt : ctx.defEPS);
  const epfER = (marked === 0 || ctx.erFromEmp !== 'Yes' || e.epsApp !== 'Yes') ? 0 : (e.erAmt > 0 ? e.erAmt : ctx.defEPSEr);
  const esi = (marked === 0 || ctx.erFromEmp !== 'Yes') ? 0 : e.esi;
  const other = marked === 0 ? 0 : e.other;
  const gross = earned + N_(ayVal);
  const totDed = absence + epfEE + epfER + other + esi;
  const netPre = gross - totDed;
  // ---- v2: salary advance / loan recovery (oldest first; never makes net salary negative) ----
  const exitThisMonth = exited && exit >= ms && exit <= me;
  const ym = s => { const dt = new Date(Date.UTC(1899, 11, 30) + s * 86400000); return dt.getUTCFullYear() * 12 + dt.getUTCMonth(); };
  const thisYM = y * 12 + (m - 1);
  let advTake = 0, advBal = 0, advChanged = false, hasAdvNow = false;
  let cap = round_(Math.max(0, netPre) * ctx.advCap, 0);           // never more than the allowed % of the net salary
  (ctx.adv[id] || []).forEach(a => {
    const old = a.rec[m - 1] || 0;
    const prior = a.rec.slice(0, m - 1).reduce((x, q) => x + (Number(q) || 0), 0);
    const remaining = round_(a.amount - a.already - prior, 2);
    let due = 0;
    if (remaining > 0 && a.date <= me && marked > 0) {
      if (exitThisMonth) due = remaining;                                            // final settlement: recover everything
      else if (ym(a.from) <= thisYM) due = a.mode === 'One-time' ? remaining : Math.min(a.inst > 0 ? a.inst : remaining, remaining);
    }
    const take = round_(Math.min(due, cap), 2);
    cap = round_(cap - take, 2);
    a.rec[m - 1] = take;
    if (Math.abs(take - old) > 0.004) advChanged = true;
    advTake += take;
    if (a.date <= me) { hasAdvNow = true; advBal += Math.max(0, round_(remaining - take, 2)); }
  });
  advTake = round_(advTake, 2);
  const shiftStartIdx = shiftIdx_(ctx, shiftOn_(ctx, id, ms)) + 1;
  const chgIn = (ctx.chg[id] || []).filter(c => c.date > ms && c.date <= me).length;
  const incIn = (ctx.inc[id] || []).filter(c => c.date > ms && c.date <= me).length;
  // ---- live items ----
  put('AR', missing);
  put('BP', (inMonth && onPay(tdy) && type[tdy - ms] === 'Work') ? 1 : 0);
  const workingStatus = e.status === 'Active' || exited;
  let rem = eligCal === 0 ? (exit && exit < ms ? `${e.status} on ${fmtSerial_(exit)} — not on payroll` : 'Not employed in this month (joins later)')
    : missing > 0 ? `MISSING (${missing}): ${missDates.join(', ')}`
    : marked === 0 ? (workingStatus ? 'Attendance not entered for this period' : 'Inactive employee — no attendance')
    : (workingStatus ? 'OK' : 'OK (employee now inactive)');
  if (exitThisMonth) rem += `  • ${e.status.toUpperCase()} — last working day ${fmtSerial_(exit)} (final settlement)`;
  if (exited && !exit) rem += '  • ⚠ Enter the Last Working Day in Settings (column AA)';
  if (swc > 0) rem += `  • Sandwich: ${swc} day(s) deducted`;
  if (chgIn > 0) rem += '  • Shift changed this month';
  if (incIn > 0) rem += '  • Salary revised this month';
  if (advTake > 0) rem += `  • Advance recovered ${ctx.sym}${fmtINR_(advTake)} (balance ${ctx.sym}${fmtINR_(advBal)})`;
  if (exitThisMonth && advBal > 0) rem += `  • ⚠ ${ctx.sym}${fmtINR_(advBal)} advance could not be recovered`;
  put('BF', rem);
  put('AM', cnt[0]); put('AN', cnt[1]); put('AO', cnt[2]); put('AP', cnt[3]); put('AQ', cnt[4]);
  put('AS', eligWD); put('AT', marked === 0 ? 0 : (is30 ? eligCal : eligWD) - unpaid); put('AU', unpaid);
  put('AV', marked === 0 ? '' : cnt[0] / marked); put('AW', marked === 0 ? '' : paidW / marked); put('AX', earned);
  put('AZ', gross); put('BA', absence); put('BB', epfEE + epfER); put('BC', other);
  put('BD', round_(totDed + advTake, 2)); put('BE', round_(netPre - advTake, 2));
  put('BG', join); put('BH', method); put('BI', e.status || 'Active'); put('BJ', eligCal); put('BK', eligWD); put('BL', marked);
  put('BM', paidW); put('BN', shiftStartIdx); put('BO', shiftWD); put('BR', epfEE); put('BS', epfER); put('BT', chgIn); put('BU', incIn);
  for (let d = 1; d <= 31; d++) { out[colN_('BV') - OUT_FIRST + d - 1] = type[d - 1]; out[colN_('DA') - OUT_FIRST + d - 1] = rate[d - 1]; }
  put('EF', esi); put('EG', swc); put('EH', swm); put('EI', brw); put('EJ', brm);
  flags.forEach((x, d) => { out[SWB.FLAG1 - OUT_FIRST + d] = x; });
  putN(XC.ADV, advTake); putN(XC.NETPRE, netPre); putN(XC.EXIT, exit || ''); putN(XC.ADVBAL, hasAdvNow ? advBal : '');
  return { bg: [e.name, e.dept, e.desig, shiftOn_(ctx, id, me), F, daily], am: out, advChanged: advChanged };
}
function fmtSerial_(s) { const d = new Date(Date.UTC(1899, 11, 30) + s * 86400000); return String(d.getUTCDate()).padStart(2, '0') + '-' + MONTHS[d.getUTCMonth()].slice(0, 3) + '-' + d.getUTCFullYear(); }
function fmtINR_(n) {                                    // Indian digit grouping: 1,23,456
  const s = String(Math.round(Math.abs(n))); let last3 = s.slice(-3); const rest = s.slice(0, -3);
  if (rest) last3 = ',' + last3;
  return (n < 0 ? '-' : '') + rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + last3;
}
function N_(v) { const n = Number(v); return (v === '' || v === null || isNaN(n)) ? 0 : n; }
function round_(x, dp) { const p = Math.pow(10, dp); const y = Math.round(Math.abs(x) * p + 1e-7) / p; return x < 0 ? -y : y; }

/* ========================== AUTOMATIC TRIGGERS ============================= */
/** Installable edit trigger (created by setup): keeps every figure up to date after each edit. */
function onEditHandler(e) {
  if (!e || !e.range) return;
  const props = PropertiesService.getDocumentProperties().getProperties();
  if (props.SETUP_STEP !== undefined || props.CAP_STEP !== undefined) return;           // setup / capacity change still running
  if (props.V2 !== '1') { toast_('Run Payroll System ▸ ⬆️ Upgrade to v2 once — figures will update after that.', 'Payroll', 10); return; }
  try { onEditHandlerInner_(e); }
  catch (err) {
    try { recalcAll_(); }                              // v2.2: retry once with a full recalculation
    catch (err2) { toast_('Could not recalculate automatically: ' + err2.message + ' — run Payroll System ▸ Run Self-Test to find the cause.', 'Payroll', 15); }
  }
}
function onEditHandlerInner_(e) {
  const sh = e.range.getSheet(), name = sh.getName();
  const r = e.range.getRow(), c = e.range.getColumn();
  const lastR = r + e.range.getNumRows() - 1, lastC = c + e.range.getNumColumns() - 1;
  const hit = (r1, r2, c1, c2) => lastR >= r1 && r <= r2 && lastC >= c1 && c <= c2;
  const mi = MONTHS.indexOf(name);
  if (mi >= 0) {
    // FAST PATH (v1.3): recalculate ONLY the edited employee(s); a neighbouring month is touched only when the
    // edit is near the start / end of the month (where a sandwich can cross the month boundary)
    const att = hit(APP.M_FIRST, APP.M_LAST, 8, 38), oth = hit(APP.M_FIRST, APP.M_LAST, 51, 51);
    if (!att && !oth) return;
    const r1 = Math.max(r, APP.M_FIRST), r2 = Math.min(lastR, APP.M_LAST);
    const ids = sh.getRange(r1, 1, r2 - r1 + 1, 1).getValues().map(x => String(x[0]).trim()).filter(String);
    if (!ids.length) return;
    const months = [mi + 1];
    if (att) {
      const d1 = Math.max(c, 8) - 7, d2 = Math.min(lastC, 38) - 7;
      if (d1 <= 10 && mi > 0) months.unshift(mi);
      if (d2 >= 22 && mi < 11) months.push(mi + 2);
    }
    recalcMonths_(months, ids);
    return;
  }
  if (name === REG) { if (hit(2, 2, 3, 4)) renderRegister_(); return; }
  if (name !== 'Settings') return;
  if (hit(APP.MASTER_FIRST, APP.MASTER_LAST, 1, APP.EXIT_COL)) syncEmployees();
  if (hit(12, 18, 7, 10) || hit(5, 4 + APP.HOL_ROWS, APP.HOL_COL, APP.HOL_COL + 2) || hit(8, 8, 3, 4) || hit(5, 8, 6, 6) ||
      hit(APP.CHG_FIRST, APP.CHG_LAST, 1, 12) || hit(APP.MASTER_FIRST, APP.MASTER_LAST, 1, APP.EXIT_COL)) markOffDays();
  if (hit(9, 9, 3, 4)) applyCurrencyFormats();
  if (hit(5, 9, 12, 12) || hit(13, 13, 3, 4)) refreshAttendanceFormatting();
  recalcAll_();
  if (hit(APP.MASTER_FIRST, APP.MASTER_LAST, 1, APP.EXIT_COL)) renderRegister_();
}

function installTriggers_() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(t => { if (['onEditHandler', 'dailyRefresh'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onEditHandler').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('dailyRefresh').timeBased().atHour(1).everyDays(1).create();
}
/** Menu: re-create the automatic triggers (use if figures stop updating after an edit). */
function enableAutoCalculation() { installTriggers_(); recalculatePayrollNow(); toast_('Automatic calculation is ON ✔', 'Payroll', 5); }
/** Runs every night (1 AM): refreshes MISSING counts, Remarks and "today" figures for the new day, and the Register. */
function dailyRefresh() { const pp = PropertiesService.getDocumentProperties(); if (pp.getProperty('V2') !== '1' || pp.getProperty('CAP_STEP') !== null) return; syncEmployees(); markOffDays(); recalcAll_(); renderRegister_(); }

function isLocked_(sh) { return String(sh.getRange('AG1').getValue()).indexOf('🔒') === 0; }
function monthFromText_(t) {
  const x = String(t).trim().toLowerCase(); if (x.length < 3) return null;
  return MONTHS.find(m => m.slice(0, 3).toLowerCase() === x.slice(0, 3)) || null;   // "Sep", "sept", "September 2026" all work
}

/* ================================ DASHBOARD ================================ */
function buildDashboard_(sh) {
  sh.clear();
  ensureSize_(sh, Math.max(180, 76 + APP.ROWS + 5), 26);
  sh.setTabColor(APP.RED);
  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 16);
  sh.setColumnWidths(2, 20, 104);   // B..U
  sh.setColumnWidth(22, 104); sh.setColumnWidth(23, 104);
  sh.setColumnWidth(24, 100);       // X helper
  const MLIST = MONTHS.map(m => '"' + m + '"').join(',');

  // helper cells (column X, hidden)
  sh.getRange('X1:X5').setFormulas([['="helper"'], [`=CHOOSE(MONTH(TODAY()),${MLIST})`], ['=YEAR(TODAY())=AttYear'],
  ['=IF(X3,SUM(F6,H6,J6,L6,N6),0)'], ['=IF(N15="Current Month",CurMonthName,N15)']]);

  // Title
  sh.getRange('B1:Q1').merge().setFormula('=CompanyName&"  —  HR ATTENDANCE & PAYROLL DASHBOARD"');
  styleBanner_(sh.getRange('B1:Q1'), 16); sh.setRowHeight(1, 44);
  sh.getRange('B2:Q2').merge().setFormula('="Attendance Year: "&AttYear&"     •     Today: "&TEXT(TODAY(),"dddd, dd mmm yyyy")&"     •     Updates automatically — no manual entry needed on this tab"')
    .setFontColor(APP.MUTED).setFontSize(9).setHorizontalAlignment('center');

  // TODAY
  section_(sh.getRange('B4:Q4'), "TODAY'S ATTENDANCE");
  const tI = s => `INDIRECT("'"&$X$2&"'!${s}")`;
  const cntT = k => `=IF($X$3,COUNTIF(${tI(`BQ14:BQ${APP.M_LAST}`)},INDEX(StatusNames,${k})),"—")`;
  card_(sh, 'B5:C5', 'B6:C6', "Today's Date", '=TODAY()', 'dd-mmm-yyyy', APP.NAVY);
  card_(sh, 'D5:E5', 'D6:E6', 'Active Employees', '=COUNTIFS(Emp_ID,"<>",Emp_Status,"<>Inactive")', '0', APP.NAVY);
  card_(sh, 'F5:G5', 'F6:G6', '=INDEX(StatusNames,1)&" Today"', cntT(1), '0', STATUS_COLORS[0][1]);
  card_(sh, 'H5:I5', 'H6:I6', '=INDEX(StatusNames,2)&" Today"', cntT(2), '0', STATUS_COLORS[1][1]);
  card_(sh, 'J5:K5', 'J6:K6', '=INDEX(StatusNames,3)&" Today"', cntT(3), '0', STATUS_COLORS[2][1]);
  card_(sh, 'L5:M5', 'L6:M6', '=INDEX(StatusNames,4)&" Today"', cntT(4), '0', STATUS_COLORS[3][1]);
  card_(sh, 'N5:O5', 'N6:O6', '=INDEX(StatusNames,5)&" Today"', cntT(5), '0', STATUS_COLORS[4][1]);
  card_(sh, 'P5:Q5', 'P6:Q6', 'Missing (not marked yet)', `=IF($X$3,SUMPRODUCT((${tI(`BP14:BP${APP.M_LAST}`)}=1)*ISNA(MATCH(${tI(`BQ14:BQ${APP.M_LAST}`)},StatusNames,0))),"—")`, '0', APP.RED);
  card_(sh, 'B7:C7', 'B8:C8', 'Raw Attendance %', '=IF(AND($X$3,$X$4>0),F6/$X$4,"—")', '0.0%', APP.NAVY);
  card_(sh, 'D7:E7', 'D8:E8', 'Paid Attendance %', '=IF(AND($X$3,$X$4>0),(F6*INDEX(StatusPct,1)+H6*INDEX(StatusPct,2)+J6*INDEX(StatusPct,3)+L6*INDEX(StatusPct,4)+N6*INDEX(StatusPct,5))/$X$4,"—")', '0.0%', APP.NAVY);
  card_(sh, 'F7:G7', 'F8:G8', 'Marked Today', '=IF($X$3,$X$4&" of "&($X$4+P6),"—")', '@', APP.NAVY);
  sh.getRange('H7:Q7').merge().setValue('Status').setBackground(APP.PALE).setFontColor(APP.MUTED).setFontSize(8).setFontWeight('bold');
  sh.getRange('H8:Q8').merge().setFormula(`=IF(NOT($X$3),"Today's date is outside the selected attendance year.",IF($X$4+N(P6)=0,"Today is not a working day for any shift (weekly off / holiday).",IF($X$4=0,"No attendance marked yet today — figures fill in as HR marks the "&$X$2&" sheet.","Live from the "&$X$2&" sheet, column "&DAY(TODAY())&".")))`)
    .setFontColor(APP.TEXT).setFontSize(10).setWrap(true).setVerticalAlignment('middle');
  box_(sh.getRange('H7:Q8'));
  sh.getRange('B9:Q9').merge().setFormula(
    '="Definitions —  Raw Attendance % = "&INDEX(StatusNames,1)&" ÷ employees marked.   Paid Attendance % = ("&' +
    [1, 2, 3, 4, 5].map(k => `INDEX(StatusNames,${k})&" × "&TEXT(INDEX(StatusPct,${k}),"0%")`).join('&" + "&') +
    '&") ÷ employees marked.   "&INDEX(StatusNames,4)&" is paid and is NOT counted as absence. Night shift counts on the date it starts."')
    .setFontSize(8).setFontColor(APP.MUTED).setWrap(true).setFontStyle('italic');
  sh.setRowHeight(9, 30);
  [5, 7].forEach(r => sh.setRowHeight(r, 22)); [6, 8].forEach(r => sh.setRowHeight(r, 38));

  // EMPLOYEE STATISTICS
  section_(sh.getRange('B11:Q11'), 'EMPLOYEE STATISTICS');
  card_(sh, 'B12:C12', 'B13:C13', 'Total Employees', '=COUNTIF(Emp_ID,"<>")', '0', APP.NAVY);
  card_(sh, 'D12:E12', 'D13:E13', 'Active Employees', '=COUNTIFS(Emp_ID,"<>",Emp_Status,"<>Inactive")', '0', APP.GREEN);
  card_(sh, 'F12:G12', 'F13:G13', 'Inactive Employees', '=COUNTIFS(Emp_ID,"<>",Emp_Status,"Inactive")', '0', APP.MUTED);
  // v1.4: Day card also counts the 3rd shift "Day (5-day)" (both are day-shift staff)
  card_(sh, 'H12:I12', 'H13:I13', '=INDEX(ShiftNames,1)&" Shift (incl. "&INDEX(ShiftNames,3)&") — active"',
    '=COUNTIFS(Emp_CurShift,INDEX(ShiftNames,1),Emp_Status,"<>Inactive",Emp_ID,"<>")+IF(INDEX(ShiftNames,3)="",0,COUNTIFS(Emp_CurShift,INDEX(ShiftNames,3),Emp_Status,"<>Inactive",Emp_ID,"<>"))', '0', APP.NAVY);
  card_(sh, 'J12:K12', 'J13:K13', '=INDEX(ShiftNames,2)&" Shift (active)"', '=COUNTIFS(Emp_CurShift,INDEX(ShiftNames,2),Emp_Status,"<>Inactive",Emp_ID,"<>")', '0', APP.NAVY);
  card_(sh, 'L12:M12', 'L13:M13', 'Gross Payroll (Year)', '=L33', 'money', APP.NAVY);
  card_(sh, 'N12:O12', 'N13:O13', 'Net Payroll (Year)', '=Q33', 'money', APP.RED);
  card_(sh, 'P12:Q12', 'P13:Q13', 'Missing Entries (Year)', '=I33', '0', APP.RED);
  sh.setRowHeight(13, 38);

  // PAYROLL (selected month)
  section_(sh.getRange('B15:K15'), 'PAYROLL');
  sh.getRange('L15:M15').merge().setValue('Select Month ▸').setFontWeight('bold').setHorizontalAlignment('right').setBackground(APP.PALE);
  sh.getRange('N15:O15').merge().setValue('Current Month').setBackground(APP.INPUT).setFontWeight('bold').setHorizontalAlignment('center')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Current Month'].concat(MONTHS), true).setAllowInvalid(false).build());
  sh.getRange('P15:Q15').merge().setFormula('="Showing: "&$X$5&" "&AttYear').setFontColor(APP.MUTED).setFontSize(9);
  const pm = col => `=IFERROR(INDEX(${col}21:${col}32,MATCH($X$5,$B$21:$B$32,0)),0)`;
  card_(sh, 'B16:C16', 'B17:C17', 'Payroll Month', '=$X$5&" "&AttYear', '@', APP.NAVY);
  card_(sh, 'D16:E16', 'D17:E17', 'Gross Payroll', pm('L'), 'money', APP.NAVY);
  card_(sh, 'F16:G16', 'F17:G17', 'Absence Deduction', pm('M'), 'money', STATUS_COLORS[4][1]);
  card_(sh, 'H16:I16', 'H17:I17', 'EPF (total)', pm('N'), 'money', APP.NAVY);
  card_(sh, 'J16:K16', 'J17:K17', 'Other Deductions', pm('O'), 'money', APP.NAVY);
  card_(sh, 'L16:M16', 'L17:M17', 'Total Deductions', pm('P'), 'money', APP.NAVY);
  card_(sh, 'N16:O16', 'N17:O17', 'NET PAYROLL', pm('Q'), 'money', APP.RED);
  card_(sh, 'P16:Q16', 'P17:Q17', 'Employees Paid', `=IFERROR(COUNTIF(INDIRECT("'"&$X$5&"'!BL14:BL${APP.M_LAST}"),">0"),0)`, '0', APP.NAVY);
  sh.setRowHeight(17, 38);

  // MONTHLY OVERVIEW
  section_(sh.getRange('B19:Q19'), 'MONTHLY OVERVIEW  (January → December)');
  header_(sh.getRange('B20:Q20'), ['Month', 'Active Emp.', '=INDEX(StatusNames,1)', '=INDEX(StatusNames,2)', '=INDEX(StatusNames,3)',
    '=INDEX(StatusNames,4)', '=INDEX(StatusNames,5)', 'Missing', 'Attendance %', 'Paid Attendance %', 'Gross Payroll',
    'Absence Deduction', 'EPF', 'Other Deductions', 'Total Deductions', 'Net Payroll']);
  const ov = MONTHS.map(m => {
    const q = `'${m}'!`;
    return [m, `=${q}A4`, `=${q}C4`, `=${q}D4`, `=${q}E4`, `=${q}F4`, `=${q}G4`, `=${q}N4`, `=IF(${q}H4="","",${q}H4)`, `=IF(${q}K4="","",${q}K4)`,
      `=${q}Q4`, `=${q}T4`, `=${q}W4`, `=${q}Z4`, `=${q}T4+${q}W4+${q}Z4`, `=${q}AC4`];
  });
  sh.getRange('B21:Q32').setValues(ov.map(r => [r[0]].concat(new Array(15).fill(''))));
  sh.getRange('C21:Q32').setFormulas(ov.map(r => r.slice(1)));
  const tot = ['TOTAL (Year)', ''];
  ['D', 'E', 'F', 'G', 'H', 'I'].forEach(c => tot.push(`=SUM(${c}21:${c}32)`));
  tot.push('=IFERROR(SUM(D21:D32)/SUM(D21:H32),0)');
  tot.push('=IFERROR((SUM(D21:D32)*INDEX(StatusPct,1)+SUM(E21:E32)*INDEX(StatusPct,2)+SUM(F21:F32)*INDEX(StatusPct,3)+SUM(G21:G32)*INDEX(StatusPct,4)+SUM(H21:H32)*INDEX(StatusPct,5))/SUM(D21:H32),0)');
  ['L', 'M', 'N', 'O', 'P', 'Q'].forEach(c => tot.push(`=SUM(${c}21:${c}32)`));
  sh.getRange('B33').setValue(tot[0]);
  sh.getRange('C33:Q33').setFormulas([tot.slice(1).map(x => x === '' ? '=""' : x)]);
  sh.getRange('B21:B33').setFontWeight('bold');
  sh.getRange('C21:I33').setNumberFormat('0.##').setHorizontalAlignment('center');
  sh.getRange('J21:K33').setNumberFormat('0.0%').setHorizontalAlignment('center');
  sh.getRange('B33:Q33').setBackground(APP.PALE).setFontWeight('bold');
  sh.getRange('Q21:Q33').setFontWeight('bold').setFontColor(APP.RED);
  sh.getRange('B21:Q32').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  box_(sh.getRange('B20:Q33'));
  // highlight current payroll month
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B21=$X$5').setBackground('#FFF1CC').setBold(true)
      .setRanges([sh.getRange('B21:Q32')]).build()
  ]);

  // CHART DATA (right side, visible, auto)
  section_(sh.getRange('S19:W19'), 'CHART DATA (auto — do not edit)');
  header_(sh.getRange('S20:T20'), ['Status', 'Days (selected month)']);
  const dist = [1, 2, 3, 4, 5].map(k => [`=INDEX(StatusNames,${k})`, `=IFERROR(INDEX($D$21:$H$32,MATCH($X$5,$B$21:$B$32,0),${k}),0)`]);
  sh.getRange('S21:T25').setFormulas(dist);
  header_(sh.getRange('S27:W27'), ['Shift', 'Present Days', 'Marked Days', 'Attendance %', 'Paid Attendance %']);
  const sI = c => `INDIRECT("'"&$X$5&"'!${c}14:${c}${APP.M_LAST}")`;
  const sh5 = [];
  for (let k = 1; k <= 4; k++) {
    const r = 27 + k;
    sh5.push([`=INDEX(ShiftNames,${k})&""`,
      `=IF(S${r}="","",SUMIFS(${sI('AM')},${sI('E')},S${r}))`,
      `=IF(S${r}="","",SUMIFS(${sI('BL')},${sI('E')},S${r}))`,
      `=IF(S${r}="","",IFERROR(T${r}/U${r},0))`,
      `=IF(S${r}="","",IFERROR(SUMIFS(${sI('BM')},${sI('E')},S${r})/U${r},0))`]);
  }
  sh.getRange('S28:W31').setFormulas(sh5);
  sh.getRange('V28:W31').setNumberFormat('0.0%');
  box_(sh.getRange('S20:T25')); box_(sh.getRange('S27:W31'));

  // YEARLY EMPLOYEE SUMMARY (rows 74+)
  const YF = 76, YL = YF + APP.ROWS - 1;
  section_(sh.getRange('B74:U74'), 'YEARLY EMPLOYEE SUMMARY');
  header_(sh.getRange('B75:U75'), ['Employee ID', 'Employee Name', 'Department', 'Designation', 'Shift', 'Status',
    '=INDEX(StatusNames,1)', '=INDEX(StatusNames,2)', '=INDEX(StatusNames,3)', '=INDEX(StatusNames,4)', '=INDEX(StatusNames,5)',
    'Paid Days', 'Unpaid Days', 'Attendance %', 'Paid Attendance %', 'Annual Salary (Gross)', 'Absence Deduction',
    'Total EPF', 'Total Other Deductions', 'Net Annual Salary']);
  sh.setRowHeight(75, 40); sh.getRange('B75:U75').setWrap(true);
  sh.getRange('B' + YF).setFormula('=IFERROR(FILTER(Emp_ID,Emp_ID<>""),"")');
  const yrows = [];
  const sumAll = (col, r) => MONTHS.map(m => `SUMIF('${m}'!$A$14:$A$${APP.M_LAST},$B${r},'${m}'!$${col}$14:$${col}$${APP.M_LAST})`).join('+');
  for (let r = YF; r <= YL; r++) {
    const B = `$B${r}`, Z = f => `=IF(${B}="","",${f})`;
    const lk = nm => `IFERROR(INDEX(${nm},MATCH(${B},Emp_ID,0))&"","")`;
    yrows.push([
      Z(lk('Emp_Name')), Z(lk('Emp_Dept')), Z(lk('Emp_Desig')), Z(lk('Emp_CurShift')),
      Z(`IF(${lk('Emp_Status')}="","Active",${lk('Emp_Status')})`),
      Z(sumAll('AM', r)), Z(sumAll('AN', r)), Z(sumAll('AO', r)), Z(sumAll('AP', r)), Z(sumAll('AQ', r)),
      Z(sumAll('AT', r)), Z(sumAll('AU', r)),
      Z(`IF(SUM($H${r}:$L${r})=0,"",$H${r}/SUM($H${r}:$L${r}))`),
      Z(`IF(SUM($H${r}:$L${r})=0,"",(${sumAll('BM', r)})/SUM($H${r}:$L${r}))`),
      Z(sumAll('AZ', r)), Z(sumAll('BA', r)), Z(sumAll('BB', r)), Z(sumAll('BC', r)), Z(sumAll('BE', r))
    ]);
  }
  sh.getRange(`C${YF}:U${YL}`).setFormulas(yrows);
  sh.getRange(`B${YF}:B${YL}`).setFontWeight('bold').setFontColor(APP.NAVY);
  sh.getRange(`F${YF}:G${YL}`).setHorizontalAlignment('center');
  sh.getRange(`H${YF}:N${YL}`).setNumberFormat('0.##').setHorizontalAlignment('center');
  sh.getRange(`O${YF}:P${YL}`).setNumberFormat('0.0%').setHorizontalAlignment('center');
  sh.getRange(`U${YF}:U${YL}`).setFontWeight('bold').setFontColor(APP.RED);
  sh.getRange(`B${YF}:U${YL}`).setBorder(true, true, true, true, true, true, '#E2E6EE', SpreadsheetApp.BorderStyle.SOLID);

  // CHARTS
  section_(sh.getRange('B35:Q35'), 'CHARTS  (update automatically)');
  const base = { legend: { position: 'bottom' }, backgroundColor: '#FFFFFF', titleTextStyle: { color: APP.NAVY, fontSize: 14, bold: true } };
  const addChart = (type, ranges, row, col, opts) => {
    let b = sh.newChart().setChartType(type);
    ranges.forEach(r => { b = b.addRange(sh.getRange(r)); });
    b = b.setNumHeaders(1).setPosition(row, col, 6, 6).setOption('width', 820).setOption('height', 360);
    Object.keys(base).forEach(k => { b = b.setOption(k, base[k]); });
    Object.keys(opts).forEach(k => { b = b.setOption(k, opts[k]); });
    sh.insertChart(b.build());
  };
  addChart(Charts.ChartType.LINE, ['B20:B32', 'J20:K32'], 36, 2, {
    title: 'Monthly Attendance Trend (Jan → Dec)', colors: [APP.NAVY, APP.RED], pointSize: 6, curveType: 'function',
    vAxis: { format: 'percent', viewWindow: { min: 0, max: 1 } }
  });
  addChart(Charts.ChartType.PIE, ['S20:T25'], 36, 10, {
    title: 'Attendance Distribution — Selected Payroll Month', pieHole: 0.45,
    colors: STATUS_COLORS.map(c => c[1]), legend: { position: 'right' }
  });
  addChart(Charts.ChartType.COLUMN, ['S27:S31', 'V27:W31'], 55, 2, {
    title: 'Shift-wise Attendance — Selected Payroll Month', colors: [APP.NAVY, APP.RED],
    vAxis: { format: 'percent', viewWindow: { min: 0, max: 1 } }
  });
  addChart(Charts.ChartType.BAR, [`C75:C${YL}`, `O75:P${YL}`], 55, 10, {
    title: 'Employee-wise Attendance % (Year)', colors: [APP.NAVY, APP.RED],
    hAxis: { format: 'percent', viewWindow: { min: 0, max: 1 } }
  });
  dashV2_(sh);
  sh.setFrozenRows(2);
  sh.hideColumns(24);
}

/* =============================== SALARY SLIP =============================== */
function buildSalarySlip_(sh) {
  sh.clear();
  ensureSize_(sh, 50, 12);
  sh.setTabColor(APP.GREEN);
  sh.setHiddenGridlines(true);
  [14, 185, 160, 24, 185, 160, 14, 270, 20, 20, 150].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const MLIST = MONTHS.map(m => '"' + m + '"').join(',');

  // ---- control panel (rows 1-3, not printed) ----
  sh.getRange('A1:G1').merge().setValue('SALARY SLIP GENERATOR  —  select an employee and a month; the slip below fills in automatically');
  styleBanner_(sh.getRange('A1:G1'), 11); sh.setRowHeight(1, 32);
  sh.getRange('B2').setValue('Select Employee ▸'); sh.getRange('E2').setValue('Select Month ▸');
  sh.getRange('B3').setValue('Salary Year (from Settings)'); sh.getRange('E3').setValue('Status');
  sh.getRange('B2:B3').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('E2:E3').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('C2:D2').merge().setBackground(APP.INPUT).setFontWeight('bold')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(named_(sh.getParent(), 'ActiveList'), true).setAllowInvalid(false).build());
  sh.getRange('F2').setBackground(APP.INPUT).setFontWeight('bold')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(MONTHS, true).setAllowInvalid(false).build());
  sh.getRange('C3:D3').merge().setFormula('=AttYear').setNumberFormat('0').setHorizontalAlignment('left').setFontWeight('bold');
  sh.getRange('F3').setFormula('=IF($K$6="OK","✔ Ready to print / send",$K$6)').setFontSize(9).setWrap(true);
  box_(sh.getRange('B2:F3'));
  sh.setRowHeights(2, 2, 28);
  sh.getRange('H1').setValue('Actions (click a button)').setFontWeight('bold').setFontColor(APP.MUTED).setFontSize(9);

  // ---- helper cells (column K, hidden) ----
  sh.getRange('K1:K8').setFormulas([
    ['="helper"'],
    ['=IF(C2="","",IFERROR(TRIM(LEFT(C2,FIND(" - ",C2)-1)),TRIM(C2)))'],
    ['=F2&""'],
    [`=IF(OR(K2="",K3=""),0,IFERROR(MATCH(K2,INDIRECT("'"&K3&"'!A14:A${APP.M_LAST}"),0),0))`],
    [`=IF(K4=0,0,N(INDEX(INDIRECT("'"&K3&"'!BL14:BL${APP.M_LAST}"),K4)))`],
    ['=IF(C2="","Please select an employee.",IF(F2="","Please select a month.",IF(K5=0,"Attendance not entered for this period.","OK")))'],
    [`=IF(K3="",0,IFERROR(MATCH(K3,{${MLIST}},0),0))`],
    ['=IF(K2="",0,IFERROR(MATCH(K2,Emp_ID,0),0))']
  ]);

  const emp = nm => `=IF($K$8=0,"",INDEX(${nm},$K$8))`;
  const mv = col => `=IF($K$6<>"OK","",N(INDEX(INDIRECT("'"&$K$3&"'!${col}14:${col}${APP.M_LAST}"),$K$4)))`;

  // ---- printable slip rows 5-42 ----
  sh.getRange('B5:B7').merge().setFormula('=IF(LogoURL="",UPPER(LEFT(CompanyName,1)),IMAGE(LogoURL,1))')
    .setBackground(APP.RED).setFontColor(APP.WHITE).setFontSize(30).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange('C5:F5').merge().setFormula('=CompanyName').setFontSize(18).setFontWeight('bold').setFontColor(APP.NAVY);
  sh.getRange('C6:F6').merge().setFormula('=CompanyAddress').setFontColor(APP.TEXT);
  sh.getRange('C7:F7').merge().setFormula('="Phone: "&CompanyPhone&"     |     Email: "&CompanyEmail').setFontColor(APP.TEXT);
  sh.getRange('C5:F7').setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.setRowHeight(5, 34); sh.setRowHeights(6, 2, 20);
  sh.getRange('B8:F8').setBackground(APP.RED); sh.setRowHeight(8, 5);
  sh.getRange('B9:F9').merge().setValue('SALARY SLIP').setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center').setFontColor(APP.NAVY);
  sh.setRowHeight(9, 36);
  sh.getRange('B10').setValue('Salary Month'); sh.getRange('C10').setFormula('=IF($K$3="","—",$K$3)');
  sh.getRange('E10').setValue('Salary Year'); sh.getRange('F10').setFormula('=AttYear');
  sh.getRange('B10:F10').setFontWeight('bold'); sh.getRange('F10').setNumberFormat('0').setHorizontalAlignment('left');
  sh.getRange('C10').setFontColor(APP.RED); sh.getRange('F10').setFontColor(APP.RED);
  sh.getRange('B11:F11').merge().setFormula('=IF($K$6="OK","",$K$6)').setFontColor(APP.RED).setFontWeight('bold').setHorizontalAlignment('center');

  const pairs = (row, l1, f1, l2, f2) => {
    sh.getRange(row, 2).setValue(l1).setFontColor(APP.MUTED);
    if (String(l1).charAt(0) === '=') sh.getRange(row, 2).setFormula(l1);
    sh.getRange(row, 3).setFormula(f1).setFontWeight('bold').setHorizontalAlignment('left');
    sh.getRange(row, 5).setValue(l2).setFontColor(APP.MUTED);
    if (String(l2).charAt(0) === '=') sh.getRange(row, 5).setFormula(l2);
    sh.getRange(row, 6).setFormula(f2).setFontWeight('bold').setHorizontalAlignment('left');
  };
  section_(sh.getRange('B12:F12'), 'EMPLOYEE DETAILS');
  pairs(13, 'Employee ID', emp('Emp_ID'), 'Employee Name', emp('Emp_Name'));
  pairs(14, 'Department', emp('Emp_Dept'), 'Designation', emp('Emp_Desig'));
  pairs(15, 'Shift', `=IF($K$8=0,"",IF($K$4>0,INDEX(INDIRECT("'"&$K$3&"'!E14:E${APP.M_LAST}"),$K$4),INDEX(Emp_CurShift,$K$8)))`, 'Joining Date', emp('Emp_Join'));
  pairs(16, 'Shift Timing',
    '=IF($C$15="","",IFERROR(TEXT(INDEX(ShiftStart,MATCH($C$15,ShiftNames,0)),"h:mm AM/PM")&" – "&TEXT(INDEX(ShiftEnd,MATCH($C$15,ShiftNames,0)),"h:mm AM/PM")&IF(INDEX(ShiftCross,MATCH($C$15,ShiftNames,0))="Yes"," (next day)",""),""))',
    'Pay Period', '=IF($K$7=0,"",TEXT(DATE(AttYear,$K$7,1),"dd-mmm-yyyy")&" to "&TEXT(EOMONTH(DATE(AttYear,$K$7,1),0),"dd-mmm-yyyy"))');
  sh.getRange('F15').setNumberFormat('dd-mmm-yyyy');
  box_(sh.getRange('B13:F16'));

  // ---- earnings (salary heads) & deductions: rows 18-28 ----
  // K9 = pro-rata factor (earned ÷ rate; 1 for a normal full month), used to split earned salary across heads
  sh.getRange('K9').setFormula(`=IF($K$6<>"OK",0,IFERROR(INDEX(INDIRECT("'"&$K$3&"'!AX14:AX${APP.M_LAST}"),$K$4)/INDEX(INDIRECT("'"&$K$3&"'!F14:F${APP.M_LAST}"),$K$4),0))`);
  header_(sh.getRange('B18:C18'), ['EARNINGS', 'AMOUNT']);
  header_(sh.getRange('E18:F18'), ['DEDUCTIONS', 'AMOUNT']);
  sh.getRange('C18').setHorizontalAlignment('right'); sh.getRange('F18').setHorizontalAlignment('right');
  const hd = i => `=IF($K$6<>"OK","",ROUND(INDEX(Emp_H${i},$K$8)*$K$9,0))`;
  const earn = [
    ['=INDEX(HeadNames,1,1)', `=IF($K$6<>"OK","",INDEX(INDIRECT("'"&$K$3&"'!AX14:AX${APP.M_LAST}"),$K$4)-SUM(C20:C26))`],   // Basic absorbs rounding
    ['=INDEX(HeadNames,1,2)', hd(2)], ['=INDEX(HeadNames,1,3)', hd(3)], ['=INDEX(HeadNames,1,4)', hd(4)],
    ['=INDEX(HeadNames,1,5)', hd(5)], ['=INDEX(HeadNames,1,6)', hd(6)], ['=INDEX(HeadNames,1,7)', hd(7)],
    ['Increment / Salary Revision', `=IF($K$6<>"OK","",ROUND((INDEX(INDIRECT("'"&$K$3&"'!F14:F${APP.M_LAST}"),$K$4)-INDEX(Emp_Salary,$K$8))*$K$9,0))`],
    ['Other Earnings', mv('AY')]
  ];
  const ded = [
    [`=IF($K$6<>"OK","LOP / Absence Deduction",IF(N(INDEX(INDIRECT("'"&$K$3&"'!EG14:EG${APP.M_LAST}"),$K$4))>0,"LOP / Absence (incl. "&INDEX(INDIRECT("'"&$K$3&"'!EG14:EG${APP.M_LAST}"),$K$4)&" sandwich day"&IF(INDEX(INDIRECT("'"&$K$3&"'!EG14:EG${APP.M_LAST}"),$K$4)=1,"","s")&")","LOP / Absence Deduction"))`, mv('BA')],
    ['EPF – Employee Share', mv('BR')],
    ['=IF(ErFromEmp="Yes","EPF – Employer Share (part of CTC)","EPF – Employer Share (paid by company)")', mv('BS')],
    ['=IF(ErFromEmp="Yes","ESI – Employer (part of CTC)","ESI – Employer (paid by company)")', mv('EF')],
    ['Other Deduction', mv('BC')]
  ];
  for (let i = 0; i < 9; i++) {
    const r = 19 + i;
    sh.getRange(r, 2).setFormula(earn[i][0].charAt(0) === '=' ? earn[i][0] : `="${earn[i][0]}"`);
    sh.getRange(r, 3).setFormula(earn[i][1]);
    if (ded[i]) {
      sh.getRange(r, 5).setFormula(ded[i][0].charAt(0) === '=' ? ded[i][0] : `="${ded[i][0]}"`);
      sh.getRange(r, 6).setFormula(ded[i][1]);
    }
  }
  sh.getRange('B28').setValue('MONTHLY GROSS SALARY'); sh.getRange('C28').setFormula(mv('AZ'));
  sh.getRange('E28').setValue('TOTAL DEDUCTIONS'); sh.getRange('F28').setFormula(mv('BD'));
  sh.getRange('C19:C28').setHorizontalAlignment('right'); sh.getRange('F19:F28').setHorizontalAlignment('right');
  sh.getRange('B28:F28').setFontWeight('bold').setBackground(APP.PALE);
  sh.getRange('D18:D28').setBackground(null);
  box_(sh.getRange('B18:C28')); box_(sh.getRange('E18:F28'));
  sh.getRange('B28:C28').setBorder(true, null, null, null, null, null, APP.NAVY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange('E28:F28').setBorder(true, null, null, null, null, null, APP.NAVY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // ---- net (rows 30-32) ----
  sh.getRange('B30:E30').merge().setValue('Monthly Gross Salary'); sh.getRange('F30').setFormula(mv('AZ'));
  sh.getRange('B31:E31').merge().setValue('Less: Total Deductions'); sh.getRange('F31').setFormula(mv('BD'));
  sh.getRange('B32:E32').merge().setValue('NET SALARY PAYABLE (IN HAND)'); sh.getRange('F32').setFormula(mv('BE'));
  sh.getRange('B30:F31').setFontWeight('bold'); sh.getRange('F30:F32').setHorizontalAlignment('right');
  sh.getRange('B32:F32').setBackground(APP.NAVY).setFontColor(APP.WHITE).setFontSize(14).setFontWeight('bold');
  sh.setRowHeight(32, 34); sh.getRange('B32:F32').setVerticalAlignment('middle');
  box_(sh.getRange('B30:F32'), APP.NAVY);
  sh.getRange('B33:F33').merge().setFormula('=IF($K$6<>"OK","","Amount in words: "&AMOUNTINWORDS(F32,CurrencySym))')
    .setFontStyle('italic').setFontColor(APP.TEXT).setWrap(true);
  sh.setRowHeight(33, 30);
  sh.getRange('B34:F34').merge().setFormula(
    '=IF($K$6<>"OK","",IF(N(INDEX(INDIRECT("\'"&$K$3&"\'!BU14:BU' + APP.M_LAST + '"),$K$4))>0,"Salary revised during this month — earnings are pro-rated day-wise at the old and new rates. ","")&IF(N(INDEX(INDIRECT("\'"&$K$3&"\'!BT14:BT' + APP.M_LAST + '"),$K$4))>0,"Shift changed during this month (shift shown is at month end). ",""))&' +
    '"Earnings are pro-rated for mid-month joining or salary revision. Weekly offs & holidays are paid"&IF(SandwichOn="Yes",", except sandwich days (weekly off / holiday between Leave or Absent days), which are deducted","")&"."')
    .setFontSize(8).setFontColor(APP.MUTED).setWrap(true);
  sh.setRowHeight(34, 40);
  slipV2_(sh);

  sh.getRange('B38:C38').merge().setValue('Employee Signature: ______________________');
  sh.getRange('E38:F38').merge().setValue('Authorized By: ______________________').setHorizontalAlignment('right');
  sh.getRange('B38:F38').setFontWeight('bold');
  sh.getRange('B40:F40').merge().setFormula('="This is a computer-generated salary slip.   "&CompanyName&"  •  "&CompanyEmail&"  •  "&CompanyPhone')
    .setFontSize(8).setFontColor(APP.MUTED).setHorizontalAlignment('center');
  sh.getRange('B40:F40').setBorder(true, null, null, null, null, null, APP.RED, SpreadsheetApp.BorderStyle.SOLID);
  box_(sh.getRange('A4:G41'), APP.NAVY);   // outer frame of the printed page
  sh.getParent().setNamedRange('SlipPrintArea', sh.getRange('A4:G41'));

  // status colour
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('✔').setFontColor(APP.GREEN).setBold(true).setRanges([sh.getRange('F3')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$K$6<>"OK"').setFontColor(APP.RED).setBold(true).setRanges([sh.getRange('F3')]).build()
  ]);

  // default selection: first sample employee, current payroll month
  sh.getRange('C2').setValue('EMP001 - Employee 1');
  sh.getRange('F2').setValue(MONTHS[Number(Utilities.formatDate(new Date(), 'Asia/Kolkata', 'M')) - 1]);

  // buttons (clickable images)
  try {
    const b1 = sh.insertImage(Utilities.newBlob(Utilities.base64Decode(BTN_PDF_B64), 'image/png', 'pdf.png'), 8, 2, 8, 2);
    b1.assignScript('generateSalarySlipPDF'); b1.setAltTextTitle('Generate Salary Slip PDF');
    const b2 = sh.insertImage(Utilities.newBlob(Utilities.base64Decode(BTN_MAIL_B64), 'image/png', 'mail.png'), 8, 3, 8, 2);
    b2.assignScript('sendSalarySlipEmail'); b2.setAltTextTitle('Send Salary Slip');
  } catch (err) { /* buttons optional — the Payroll System menu does the same */ }
  sh.hideColumns(11);
  sh.setFrozenRows(3);
}

/* ============================ PDF + EMAIL ================================== */
function generateSalarySlipPDF() { safeRun_(generateSalarySlipPDF_, 'Salary Slip PDF'); }
function sendSalarySlipEmail() { safeRun_(sendSalarySlipEmail_, 'Send Salary Slip'); }
function downloadRegisterPDF() { safeRun_(downloadRegisterPDF_, 'Payroll Register PDF'); }
function generateSalarySlipPDF_() {
  const info = slipInfo_(); if (!info) return;
  const blob = exportSlipPdf_(info);
  const folder = slipFolder_();
  const file = folder.createFile(blob);
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial;padding:6px">
      <p style="margin:0 0 8px"><b>${esc_(file.getName())}</b> has been created.</p>
      <p style="margin:0 0 14px;color:#555">Saved in Google Drive ▸ folder "${esc_(folder.getName())}".</p>
      <a href="https://drive.google.com/uc?export=download&id=${file.getId()}" target="_blank"
         style="background:${APP.NAVY};color:#fff;padding:9px 16px;border-radius:6px;text-decoration:none;margin-right:8px">⬇ Download PDF</a>
      <a href="${file.getUrl()}" target="_blank" style="color:${APP.RED}">Open in Drive</a>
    </div>`).setWidth(430).setHeight(170);
  SpreadsheetApp.getUi().showModalDialog(html, 'Salary Slip PDF ready');
}

function sendSalarySlipEmail_() {
  const ui = SpreadsheetApp.getUi();
  const info = slipInfo_(); if (!info) return;
  if (!info.email) { ui.alert('No email address', `There is no email for ${info.name} in Settings ▸ Employee Master (column N).`, ui.ButtonSet.OK); return; }
  const ok = ui.alert('Confirm: send salary slip?',
    `Send the ${info.month} ${info.year} salary slip of ${info.name}\nto: ${info.email}\n\nNet salary shown on slip: ${info.net}\n\nThe PDF will also be saved to Drive.`,
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) { toast_('Cancelled — nothing was sent.', 'Salary Slip'); return; }
  const blob = exportSlipPdf_(info);
  slipFolder_().createFile(blob.copyBlob());
  MailApp.sendEmail({
    to: info.email,
    subject: `Salary Slip — ${info.month} ${info.year} — ${info.company}`,
    body: `Dear ${info.name},\n\nPlease find attached your salary slip for ${info.month} ${info.year}.\n\nRegards,\n${info.company}`,
    name: info.company,
    attachments: [blob]
  });
  ui.alert('Sent ✔', `Salary slip emailed to ${info.email}.`, ui.ButtonSet.OK);
}

function slipInfo_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Salary Slip');
  SpreadsheetApp.flush();
  const msg = String(sh.getRange('K6').getValue());
  if (msg !== 'OK') { SpreadsheetApp.getUi().alert('Salary Slip', msg, SpreadsheetApp.getUi().ButtonSet.OK); return null; }
  const st = ss.getSheetByName('Settings');
  const idx = Number(sh.getRange('K8').getValue());
  return {
    ss: ss, sheet: sh,
    id: String(sh.getRange('K2').getValue()),
    name: String(sh.getRange('F13').getDisplayValue()),
    month: String(sh.getRange('K3').getValue()),
    year: String(sh.getRange('F10').getDisplayValue()),
    net: String(sh.getRange('F32').getDisplayValue()),
    email: idx ? String(st.getRange(APP.MASTER_FIRST + idx - 1, 14).getValue()).trim() : '',
    company: String(st.getRange('C4').getValue())
  };
}

function exportSlipPdf_(info) {
  SpreadsheetApp.flush();
  const url = `https://docs.google.com/spreadsheets/d/${info.ss.getId()}/export?format=pdf&gid=${info.sheet.getSheetId()}` +
    '&size=A4&portrait=true&scale=4&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=UNDEFINED' +
    '&fzr=false&fzc=false&horizontal_alignment=CENTER&vertical_alignment=TOP' +
    '&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5' +
    '&r1=3&c1=0&r2=41&c2=7';                       // rows 4-41, columns A-G (the slip only)
  const resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('PDF export failed (HTTP ' + resp.getResponseCode() + '). Please try again.');
  const clean = s => s.replace(/[\\/:*?"<>|]/g, '-');
  return resp.getBlob().setName(clean(`Salary Slip - ${info.name} - ${info.month} ${info.year}.pdf`));
}

function slipFolder_(folderName) {
  const ss = SpreadsheetApp.getActive();
  const name = folderName || 'Salary Slips';
  let parent;
  try { parent = DriveApp.getFileById(ss.getId()).getParents().next(); } catch (e) { parent = DriveApp.getRootFolder(); }
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/* ============================ PAYROLL REGISTER ============================= */
/** One tab: pick a month → every employee's attendance, earnings, deductions and net payable. */
function buildRegister_(sh) {
  const RF = 9, RL = RF + APP.ROWS - 1;          // data rows 9..58 (row 8 = totals)
  sh.clear();
  ensureSize_(sh, RL + 5, 32);
  sh.setTabColor('#B7791F');
  sh.setHiddenGridlines(true);
  // A ID | B Name | C Dept | D Shift | E Rate | F Work days | G-K statuses | L Sandwich | M Missing | N Paid | O Unpaid |
  // P Earned | Q Other earn | R Gross | S LOP ded | T EPF ee | U EPF er | V ESI | W Other ded | X Advance/Loan | Y Total ded | Z NET | AA Remarks
  // helpers: AB = row in month sheet, AD2 = month
  const w = [78, 150, 105, 60, 95, 64, 56, 56, 56, 56, 56, 66, 58, 60, 62, 100, 88, 100, 95, 88, 88, 80, 88, 100, 100, 115, 260, 60, 20, 110];
  w.forEach((x, i) => sh.setColumnWidth(i + 1, x));

  sh.getRange('A1').setFormula('=CompanyName&"   —   PAYROLL REGISTER   —   "&UPPER($AD$2)&" "&AttYear');
  styleBanner_(sh.getRange('A1:AA1'), 14); sh.setRowHeight(1, 38);
  sh.getRange('A2:B2').merge().setValue('Select Month ▸').setFontWeight('bold').setHorizontalAlignment('right').setBackground(APP.PALE);
  sh.getRange('C2:D2').merge().setValue('Current Month').setBackground(APP.INPUT).setFontWeight('bold').setHorizontalAlignment('center')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Current Month'].concat(MONTHS), true).setAllowInvalid(false).build());
  sh.getRange('E2').setValue('Year').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('F2').setFormula('=AttYear').setNumberFormat('0').setFontWeight('bold');
  sh.getRange('G2:P2').merge().setFormula(`=IF(COUNTIF($A$${RF}:$A$${RL},"?*")=0,"No employees in this month.",IF(LEFT(INDIRECT("'"&$AD$2&"'!AG1"),1)="🔒","🔒 Finalized — figures are frozen","● Live — updates automatically as attendance changes"))&IF(N($M$8)>0,"     ⚠ "&$M$8&" missing attendance entr"&IF($M$8=1,"y","ies")&" — see Remarks","")`)
    .setFontColor(APP.MUTED).setFontSize(9);
  sh.getRange('AD1').setValue('helper');
  sh.getRange('AD2').setFormula('=IF($C$2="Current Month",CurMonthName,$C$2)');
  box_(sh.getRange('A2:F2'));

  card_(sh, 'A4:B4', 'A5:B5', 'Employees', `=COUNTIF($A$${RF}:$A$${RL},"?*")`, '0', APP.NAVY);
  card_(sh, 'C4:E4', 'C5:E5', 'Gross Salary', '=R8', 'money', APP.NAVY);
  card_(sh, 'F4:I4', 'F5:I5', 'LOP / Absence (incl. sandwich)', '=S8', 'money', STATUS_COLORS[4][1]);
  card_(sh, 'J4:M4', 'J5:M5', 'EPF + ESI', '=T8+U8+V8', 'money', APP.NAVY);
  card_(sh, 'N4:Q4', 'N5:Q5', 'Total Deductions (incl. advance)', '=Y8', 'money', APP.NAVY);
  card_(sh, 'R4:U4', 'R5:U5', 'NET PAYABLE', '=Z8', 'money', APP.RED);
  card_(sh, 'V4:X4', 'V5:X5', 'Missing Entries', '=M8', '0', APP.RED);
  sh.setRowHeight(5, 36);

  header_(sh.getRange('A7:AA7'), ['Employee ID', 'Employee Name', 'Department', 'Shift', 'Monthly Salary (rate)', 'Working Days',
    '=INDEX(StatusNames,1)', '=INDEX(StatusNames,2)', '=INDEX(StatusNames,3)', '=INDEX(StatusNames,4)', '=INDEX(StatusNames,5)',
    'Sandwich Days', 'Missing', 'Paid Days', 'Unpaid Days (LOP)', 'Earned Salary', 'Other Earnings', 'Gross Salary', 'LOP / Absence Deduction',
    'EPF – Employee', 'EPF – Employer', 'ESI – Employer', 'Other Deduction', 'Advance / Loan Recovery', 'Total Deductions', 'NET PAYABLE', 'Remarks']);
  sh.getRange('Z7').setBackground(APP.RED); sh.getRange('L7').setBackground('#9A3412'); sh.getRange('X7').setBackground('#5B21B6');
  sh.setRowHeight(7, 44);

  // rows are written by renderRegister_() as direct references to the selected month (fast, non-volatile)
  sh.getRange('A8:B8').merge().setValue('TOTAL');   // (no merge across the frozen A|B → C boundary)
  const tot = [];
  'EFGHIJKLMNOPQRSTUVWXYZ'.split('').forEach(c => tot.push(`=SUM(${c}${RF}:${c}${RL})`));
  tot[1] = '=""';   // working days are not summed
  sh.getRange('E8:Z8').setFormulas([tot]);
  sh.getRange('A8:AA8').setBackground('#FFF1CC').setFontWeight('bold').setBorder(true, null, true, null, null, null, APP.NAVY, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sh.getRange(`A${RF}:A${RL}`).setFontWeight('bold').setFontColor(APP.NAVY);
  sh.getRange(`D${RF}:O${RL}`).setHorizontalAlignment('center');
  sh.getRange('F8:O8').setHorizontalAlignment('center');
  sh.getRange(`F8:O${RL}`).setNumberFormat('0.##');
  sh.getRange(`Z8:Z${RL}`).setFontWeight('bold').setFontColor(APP.RED);
  sh.getRange(`X8:X${RL}`).setFontColor('#5B21B6');
  sh.getRange(`AA${RF}:AA${RL}`).setFontSize(8).setWrap(true);
  sh.getRange(`A${RF}:AA${RL}`).setBorder(true, true, true, true, true, true, '#E2E6EE', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(`A${RF}:AA${RL}`).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#FDE2E1').setFontColor('#B3261E').setBold(true)
      .setRanges([sh.getRange(`M${RF}:M${RL}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground(SANDWICH_BG).setFontColor('#9A3412').setBold(true)
      .setRanges([sh.getRange(`L${RF}:L${RL}`)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('MISSING').setFontColor('#B3261E').setRanges([sh.getRange(`AA${RF}:AA${RL}`)]).build()
  ]);
  sh.hideColumns(28, 3);           // AB..AD helpers
  sh.setFrozenRows(8);
  sh.setFrozenColumns(2);
  try {
    const b = sh.insertImage(Utilities.newBlob(Utilities.base64Decode(BTN_REG_B64), 'image/png', 'reg.png'), 27, 2, 4, 0);
    b.assignScript('downloadRegisterPDF'); b.setAltTextTitle('Download Payroll Register PDF');
  } catch (err) { /* menu item does the same */ }
}

/** Month shown on the register: the selector, or the current payroll month (same rule as Settings C14). */
function registerMonth_(ss) {
  const sel = String(ss.getSheetByName(REG).getRange('C2').getValue());
  if (MONTHS.indexOf(sel) >= 0) return sel;
  const y = Number(ss.getSheetByName('Settings').getRange('C8').getValue());
  const now = new Date(), ty = Number(Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'yyyy'));
  if (ty > y) return 'December';
  if (ty < y) return 'January';
  return MONTHS[Number(Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'M')) - 1];
}

/** Fills the register with direct cell references to the selected month's rows (called when the month or employees change). */
function renderRegister_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(REG); if (!sh) return;
  const month = registerMonth_(ss);
  const msh = ss.getSheetByName(month); if (!msh) return;
  const F = APP.M_FIRST, N = APP.ROWS, RF = 9;
  const v = msh.getRange(F, 1, N, colN_('BL')).getValues();
  const q = `'${month}'!`;
  const map = ['A', 'B', 'C', 'E', 'F', 'BO', 'AM', 'AN', 'AO', 'AP', 'AQ', 'EG', 'AR', 'AT', 'AU', 'AX', 'AY', 'AZ', 'BA', 'BR', 'BS', 'EF', 'BC', colL_(XC.ADV), 'BD', 'BE', 'BF'];
  const rows = [];
  v.forEach((x, i) => {
    const id = String(x[0]).trim();
    if (!id || !(String(x[colN_('BI') - 1]) === 'Active' || Number(x[colN_('BL') - 1]) > 0)) return;
    const r = F + i;
    rows.push(map.map(c => (c === 'AY' || c === colL_(XC.ADV)) ? `=N(${q}${c}${r})` : `=${q}${c}${r}`).concat([r]));
  });
  sh.getRange(RF, 1, N, 28).clearContent();
  if (rows.length) sh.getRange(RF, 1, rows.length, 28).setValues(rows);   // A..AA + row number in AB
  sh.getRange('AD3').setValue(month);
}

/** Saves the selected month's register as a landscape A4 PDF in Drive ▸ "Payroll Registers". */
function downloadRegisterPDF_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(REG);
  SpreadsheetApp.flush();
  const month = registerMonth_(ss), year = String(sh.getRange('F2').getDisplayValue());
  const n = Number(sh.getRange('A5').getValue()) || 0;
  if (!n) { SpreadsheetApp.getUi().alert('Payroll Register', 'No employees / attendance for ' + month + '.', SpreadsheetApp.getUi().ButtonSet.OK); return; }
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${sh.getSheetId()}` +
    '&size=A4&portrait=false&scale=2&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=CENTER' +
    '&fzr=true&fzc=false&top_margin=0.4&bottom_margin=0.4&left_margin=0.3&right_margin=0.3' +
    `&r1=0&c1=0&r2=${8 + n}&c2=27`;
  const resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('PDF export failed (HTTP ' + resp.getResponseCode() + '). Please try again.');
  const folder = slipFolder_('Payroll Registers');
  const file = folder.createFile(resp.getBlob().setName(`Payroll Register - ${month} ${year}.pdf`));
  const html = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial;padding:6px"><p style="margin:0 0 8px"><b>${esc_(file.getName())}</b> has been created.</p>
     <p style="margin:0 0 14px;color:#555">Saved in Google Drive ▸ folder "Payroll Registers".</p>
     <a href="https://drive.google.com/uc?export=download&id=${file.getId()}" target="_blank"
        style="background:${APP.NAVY};color:#fff;padding:9px 16px;border-radius:6px;text-decoration:none;margin-right:8px">⬇ Download PDF</a>
     <a href="${file.getUrl()}" target="_blank" style="color:${APP.RED}">Open in Drive</a></div>`).setWidth(430).setHeight(170);
  SpreadsheetApp.getUi().showModalDialog(html, 'Payroll Register PDF ready');
}

/**
 * Converts an amount into words (Indian numbering: lakh / crore).
 * @param {number} amount  The amount.
 * @param {string} symbol  Currency symbol (₹ gives "Rupees … Only").
 * @return {string}
 * @customfunction
 */
function AMOUNTINWORDS(amount, symbol) {
  if (amount === '' || amount === null || amount === undefined || isNaN(Number(amount))) return '';
  let n = Number(amount);
  const neg = n < 0; n = Math.abs(Math.round(n * 100) / 100);
  const rupees = Math.floor(n), paise = Math.round((n - rupees) * 100);
  const ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve',
    'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
  const two = x => x < 20 ? ones[x] : tens[Math.floor(x / 10)] + (x % 10 ? ' ' + ones[x % 10] : '');
  const three = x => { const h = Math.floor(x / 100), r = x % 100; return (h ? ones[h] + ' Hundred' + (r ? ' ' : '') : '') + (r ? two(r) : ''); };
  const indian = x => {
    if (x === 0) return '';
    const parts = [];
    const crore = Math.floor(x / 10000000); x %= 10000000;
    const lakh = Math.floor(x / 100000); x %= 100000;
    const th = Math.floor(x / 1000); x %= 1000;
    if (crore) parts.push(indian(crore) + ' Crore');
    if (lakh) parts.push(two(lakh) + ' Lakh');
    if (th) parts.push(two(th) + ' Thousand');
    if (x) parts.push(three(x));
    return parts.join(' ');
  };
  const sym = (symbol === undefined || symbol === null) ? '₹' : String(symbol).trim();
  const unit = (sym === '₹' || sym === '' || /^(INR|Rs\.?)$/i.test(sym)) ? 'Rupees ' : '';
  const words = rupees === 0 ? 'Zero' : indian(rupees);
  return (neg ? 'Minus ' : '') + unit + words + (paise ? ' and ' + two(paise) + ' Paise' : '') + ' Only';
}

/* ========================= v2 BUILDERS (also used by the Upgrade) ========================= */
/** Employee Master: Terminated / Discontinued status + Last Working Day (column AA). Safe to run again. */
function masterV2_(sh) {
  const MF = APP.MASTER_FIRST, ML = APP.MASTER_LAST, N = APP.ROWS, XCOL = APP.EXIT_COL;
  sh.setColumnWidth(XCOL, 120);
  header_(sh.getRange(23, XCOL, 1, 1), ['Last Working Day (exit date)']);
  sh.getRange(23, XCOL).setBackground('#9B1C1C').setNote(
    'Fill ONLY when Employment Status = Terminated or Discontinued.\n' +
    '• Salary is paid up to and including this date (pro-rata).\n' +
    '• The employee is removed from every later month automatically.\n' +
    '• Any advance / loan balance is recovered in full from the final salary (as far as the net salary allows).');
  sh.getRange(MF, XCOL, N, 1).setBackground(APP.INPUT).setNumberFormat('dd-mmm-yyyy').setHorizontalAlignment('center')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText('Last working day (only for Terminated / Discontinued)').build());
  sh.getRange(MF, 15, N, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive'].concat(EXITED), true).setAllowInvalid(false).build());
  sh.getRange('O23').setNote('Active = on payroll.\nInactive = temporarily off payroll (e.g. long unpaid leave).\n' +
    'Terminated / Discontinued = has left the company — ALSO enter the Last Working Day in column AA. The row stays here as a permanent record (past salary slips need it) and turns red.');
  box_(sh.getRange(23, 1, N + 1, XCOL));
  const all = sh.getRange(MF, 1, N, XCOL);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=COUNTIF($A$${MF}:$A$${ML},$A${MF})>1`)
      .setBackground('#F28B82').setRanges([sh.getRange(MF, 1, N, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(OR($O${MF}="Terminated",$O${MF}="Discontinued"),$AA${MF}="")`)
      .setBackground('#FF8A80').setFontColor('#7F0000').setRanges([sh.getRange(MF, XCOL, N, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=OR($O${MF}="Terminated",$O${MF}="Discontinued")`)
      .setBackground('#FDE2E1').setFontColor('#9B1C1C').setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=$O${MF}="Inactive"`)
      .setBackground('#E0E0E0').setFontColor('#757575').setRanges([all]).build()
  ]);
}

/** Settings ▸ I. SALARY ADVANCE & LOAN REGISTER (rows 180–232). Safe to run on an existing file. */
function buildAdvanceSection_(sh) {
  const T = APP.ADV_TITLE, F = APP.ADV_FIRST, L = APP.ADV_LAST, N = APP.ADV_ROWS;
  ensureSize_(sh, L + 5, 36);
  // Company setting (row 20): cap on advance recovery per month
  sh.getRange('A20:B20').merge().setValue('Max advance / loan recovery in a month (% of net salary)').setFontWeight('bold').setFontColor(APP.TEXT).setBackground(APP.PALE).setWrap(true);
  const capC = sh.getRange('C20:D20').merge();
  if (capC.getValue() === '') capC.setValue(1);
  capC.setNumberFormat('0%').setBackground(APP.INPUT).setHorizontalAlignment('left')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(0, 1).setAllowInvalid(false).setHelpText('0% – 100% of the net salary').build());
  sh.getRange('C20').setNote('100% = recover the full instalment if the net salary allows.\n50% = never recover more than half of the net salary in a month (the rest carries forward).\nNote: the Code on Wages / Payment of Wages Act limits TOTAL deductions to 50% of wages — check with your consultant.');
  sh.setRowHeight(20, 32);
  box_(sh.getRange('A4:D20'));
  section_(sh.getRange(T, 1, 1, 26), 'I.  SALARY ADVANCE & LOAN REGISTER  (recovered automatically from salary)');
  sh.getRange(T + 1, 1, 1, 26).merge().setValue(
    'One row per advance / loan.  ONE-TIME = the full balance is deducted from the net salary of the "Recover From" month.  ' +
    'MONTHLY = the instalment is deducted every month until the balance is zero — after that the salary returns to normal automatically.  ' +
    'If the net salary is not enough, the shortfall carries forward to the next month.  On Termination / Discontinuation the full balance is recovered from the final salary.  ' +
    'Grey columns fill in automatically — do not type in them.')
    .setFontSize(8).setFontColor(APP.MUTED).setWrap(true).setVerticalAlignment('top');
  sh.setRowHeight(T + 1, 44);
  header_(sh.getRange(T + 2, 1, 1, 26), ['Employee ID', 'Employee Name (auto)', 'Date Given', 'Type', 'Amount Given', 'Recovery Mode',
    'Monthly Instalment', 'Recover From (month)', 'Already Recovered / Repaid in Cash', 'Remarks', 'Recovered This Year (auto)',
    'Balance Outstanding (auto)', 'Status (auto)'].concat(MONTHS.map(x => x.slice(0, 3) + ' recovered (auto)')).concat(['Months Left (auto)']));
  sh.getRange(T + 2, 1, 1, 26).setBackground('#5B21B6');
  sh.setRowHeight(T + 2, 44);
  sh.getRange(F, 1, N, 1).setBackground(APP.INPUT);
  sh.getRange(F, 3, N, 8).setBackground(APP.INPUT);
  sh.getRange(F, 2, N, 1).setBackground(APP.AUTO).setFontColor(APP.NAVY);
  sh.getRange(F, 11, N, 16).setBackground(APP.AUTO);
  sh.getRange(F, 12, N, 1).setFontWeight('bold').setFontColor(APP.RED);
  const bF = [], kF = [];
  for (let r = F; r <= L; r++) {
    bF.push([`=IF(A${r}="","",IFERROR(INDEX(Emp_Name,MATCH(A${r},Emp_ID,0)),"⚠ ID not in Employee Master"))`]);
    const st = `IFERROR(INDEX(Emp_Status,MATCH(A${r},Emp_ID,0)),"")`;
    kF.push([`=IF(A${r}="","",SUM(N${r}:Y${r}))`,
      `=IF(OR(A${r}="",N(E${r})=0),"",MAX(0,N(E${r})-N(I${r})-N(K${r})))`,
      `=IF(OR(A${r}="",N(E${r})=0),"",IF(N(L${r})<=0,"✔ Settled",IF(OR(${st}="Terminated",${st}="Discontinued"),"⚠ Employee exited — balance unrecovered",IF(N(K${r})+N(I${r})=0,"Not started","Recovering"))))`]);
  }
  sh.getRange(F, 2, N, 1).setFormulas(bF);
  sh.getRange(F, 11, N, 3).setFormulas(kF);
  sh.getRange(F, 26, N, 1).setFormulas(bF.map((x, i) => { const r = F + i;
    return [`=IF(OR(A${r}="",N(L${r})<=0),"",IF(F${r}="One-time",1,IF(N(G${r})>0,ROUNDUP(L${r}/G${r},0),"—")))`]; }));
  const idRule = SpreadsheetApp.newDataValidation().requireValueInRange(sh.getRange(APP.MASTER_FIRST, 1, APP.ROWS, 1), true).setAllowInvalid(false).setHelpText('Pick an Employee ID from the Employee Master').build();
  const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  const amtRule = SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).setHelpText('Amount in rupees').build();
  sh.getRange(F, 1, N, 1).setDataValidation(idRule);
  sh.getRange(F, 3, N, 1).setDataValidation(dateRule).setNumberFormat('dd-mmm-yyyy');
  sh.getRange(F, 4, N, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Salary Advance', 'Loan'], true).setAllowInvalid(false).build());
  sh.getRange(F, 5, N, 1).setDataValidation(amtRule);
  sh.getRange(F, 6, N, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Monthly', 'One-time'], true).setAllowInvalid(false).build()).setHorizontalAlignment('center');
  sh.getRange(F, 7, N, 1).setDataValidation(amtRule);
  sh.getRange(F, 8, N, 1).setDataValidation(dateRule).setNumberFormat('mmm-yyyy').setHorizontalAlignment('center');
  sh.getRange(F, 9, N, 1).setDataValidation(amtRule);
  sh.getRange(F, 13, N, 1).setFontSize(8).setWrap(true);
  sh.getRange(F, 26, N, 1).setHorizontalAlignment('center');
  sh.getRange(T + 2, 6).setNote('Monthly = deduct the instalment (column G) every month.\nOne-time = deduct the whole balance in the "Recover From" month.');
  sh.getRange(T + 2, 7).setNote('For Monthly recovery only. Blank = whole balance in one month.');
  sh.getRange(T + 2, 8).setNote('First salary month to deduct from — any date in that month works (e.g. 01-Oct-2026). Blank = the month of the Date Given.');
  sh.getRange(T + 2, 9).setNote('Optional. Amount already recovered before this year / outside this system, or repaid by the employee in cash. Reduces the balance.');
  box_(sh.getRange(T + 2, 1, N + 1, 26), '#5B21B6');
  const stR = sh.getRange(F, 13, N, 1);
  const rules = sh.getConditionalFormatRules().filter(r => !r.getRanges().some(g => g.getRow() === F && g.getColumn() === 13));
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('✔').setFontColor(APP.GREEN).setBold(true).setRanges([stR]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('⚠').setBackground('#FDE2E1').setFontColor('#9B1C1C').setBold(true).setRanges([stR]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Recovering').setFontColor('#8A5A00').setBold(true).setRanges([stR]).build());
  sh.setConditionalFormatRules(rules);
}

/** Month sheet: FP = Advance / Loan Recovery (visible, just right of Remarks); FQ..FS hidden helpers. */
function monthV2_(sh) {
  const F = APP.M_FIRST, L = APP.M_LAST;
  ensureSize_(sh, L + 5, OUT_LAST + 3);
  sh.getRange(13, XC.ADV, 1, 4).setValues([['Advance / Loan Recovery', 'h:Net before advance', 'h:Last working day', 'h:Advance balance after month']]);
  styleHeader_(sh.getRange(13, XC.ADV, 1, 4));
  sh.getRange(13, XC.ADV).setBackground('#5B21B6').setNote('Deducted automatically from Settings ▸ I. Advance & Loan Register.\nAlready included in Total Deductions (BD) and NET SALARY (BE).');
  sh.getRange('BD13').setNote('= Absence + EPF + ESI (if CTC) + Other Deduction + Advance / Loan Recovery (column FP, right of Remarks).');
  sh.setColumnWidth(XC.ADV, 110);
  sh.getRange(F, XC.ADV, APP.ROWS, 1).setFontWeight('bold').setFontColor('#5B21B6')
    .setBorder(true, true, true, true, true, true, '#E2E6EE', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(F, XC.EXIT, APP.ROWS, 1).setNumberFormat('dd-mmm-yyyy');
  sh.showColumns(XC.ADV);
  sh.hideColumns(XC.NETPRE, 3);
}

/** Salary Slip: advance recovery line (E24/F24) and a final-settlement / balance note (row 35). */
function slipV2_(sh) {
  const L = APP.M_LAST;
  const ref = c => `INDEX(INDIRECT("'"&$K$3&"'!${c}14:${c}${L}"),$K$4)`;
  const FP = colL_(XC.ADV), FR = colL_(XC.EXIT), FS = colL_(XC.ADVBAL);
  sh.getRange('E24').setFormula('="Salary Advance / Loan Recovery"');
  sh.getRange('F24').setFormula(`=IF($K$6<>"OK","",N(${ref(FP)}))`).setHorizontalAlignment('right');
  sh.getRange('B35:F35').merge().setFormula(
    `=IF($K$6<>"OK","",IF(AND(N(${ref(FR)})>0,MONTH(N(${ref(FR)}))=$K$7,YEAR(N(${ref(FR)}))=AttYear),"FINAL SETTLEMENT — last working day "&TEXT(${ref(FR)},"dd-mmm-yyyy")&".   ","")&` +
    `IF(N(${ref(FS)})>0,"Advance / loan balance outstanding after this month: "&CurrencySym&TEXT(${ref(FS)},"#,##0"),""))`)
    .setFontSize(9).setFontColor('#5B21B6').setFontWeight('bold').setWrap(true);
  sh.setRowHeight(35, 30);
}

/** Dashboard: Terminated / Discontinued are not "active"; Total Deductions = Gross − Net (includes ESI & advance). */
function dashV2_(sh) {
  const ACT = 'Emp_ID,"<>",Emp_Status,"<>Inactive",Emp_Status,"<>Terminated",Emp_Status,"<>Discontinued"';
  sh.getRange('D6').setFormula(`=COUNTIFS(${ACT})`);
  sh.getRange('D13').setFormula(`=COUNTIFS(${ACT})`);
  sh.getRange('F12').setValue('Inactive / Exited');
  sh.getRange('F13').setFormula(`=COUNTIF(Emp_ID,"<>")-COUNTIFS(${ACT})`);
  sh.getRange('H12').setFormula('=INDEX(ShiftNames,1)&" Shift (incl. "&INDEX(ShiftNames,3)&") — active"');
  sh.getRange('H13').setFormula(`=COUNTIFS(Emp_CurShift,INDEX(ShiftNames,1),${ACT})+IF(INDEX(ShiftNames,3)="",0,COUNTIFS(Emp_CurShift,INDEX(ShiftNames,3),${ACT}))`);
  sh.getRange('J13').setFormula(`=COUNTIFS(Emp_CurShift,INDEX(ShiftNames,2),${ACT})`);
  sh.getRange('P20').setValue('Total Deductions (incl. ESI & advance)');
  sh.getRange('P21:P32').setFormulas(MONTHS.map(m => [`='${m}'!Q4-'${m}'!AC4`]));
}

/* =============================== UPGRADE ================================== */
/** Adds the v2 features to an EXISTING file without touching employees, attendance, holidays or settings. */
function upgradeToV2() {
  const ss = SpreadsheetApp.getActive(), ui = safeUi_();
  const st = ss.getSheetByName('Settings');
  if (!st || MONTHS.some(m => !ss.getSheetByName(m))) { if (ui) ui.alert('This file has not been set up yet — run setupPayrollSystem first.'); return; }
  if (ui && ui.alert('Upgrade to v2?', 'Adds:\n• Salary Advance & Loan Register (one-time / monthly recovery)\n• Terminated / Discontinued employees with Last Working Day\n\n' +
    'Your employees, attendance, holidays, salary changes and settings are KEPT.\nTakes 1–3 minutes. Continue?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    toast_('Upgrading… step 1 of 5: Settings', 'Payroll upgrade', 60);
    ensureSize_(st, APP.ADV_LAST + 5, 36);
    masterV2_(st);
    if (String(st.getRange(APP.ADV_TITLE, 1).getValue()).indexOf('I.') !== 0) buildAdvanceSection_(st);
    defineNamedRanges_(ss);
    toast_('Upgrading… step 2 of 5: month sheets', 'Payroll upgrade', 60);
    MONTHS.forEach(m => monthV2_(ss.getSheetByName(m)));
    toast_('Upgrading… step 3 of 5: Payroll Register, Dashboard, Salary Slip', 'Payroll upgrade', 60);
    const old = ss.getSheetByName(REG), sel = old.getRange('C2').getValue(), idx = old.getIndex();
    ss.deleteSheet(old);                                   // the register holds no typed data — rebuild it with the new column
    const reg = ss.insertSheet(REG, idx - 1);
    buildRegister_(reg);
    if (sel) reg.getRange('C2').setValue(sel);
    dashV2_(ss.getSheetByName('Dashboard'));
    slipV2_(ss.getSheetByName('Salary Slip'));
    PropertiesService.getDocumentProperties().setProperty('V2', '1');
    toast_('Upgrading… step 4 of 5: recalculating all months', 'Payroll upgrade', 60);
    installTriggers_();
    syncEmployees(); markOffDays();
    recalcAll_(); renderRegister_();
    toast_('Upgrading… step 5 of 5: colours, formats & protection', 'Payroll upgrade', 60);
    refreshAttendanceFormatting(); applyCurrencyFormats(); applyProtections_(ss);
  } finally { lock.releaseLock(); }
  toast_('Upgrade complete ✔', 'Payroll', 10);
  if (ui) ui.alert('Upgrade complete ✔', 'New:\n• Settings ▸ I. SALARY ADVANCE & LOAN REGISTER (row ' + APP.ADV_TITLE + ') — or menu Payroll System ▸ Advance & Loan Register.\n' +
    '• Employee Master ▸ Employment Status now has Terminated / Discontinued, plus Last Working Day in column AA.\n' +
    '• Month sheets: "Advance / Loan Recovery" column just right of Remarks.\n• Salary Slip and Payroll Register show the advance recovery.\n\n' +
    'Reload the browser tab once to see the new menu items, then run Payroll System ▸ Run Self-Test.', ui.ButtonSet.OK);
}
function menuUpgrade() {
  const ui = safeUi_();
  try { upgradeToV2(); } catch (err) { if (ui) ui.alert('Upgrade could not finish', err.message + '\n\nRun it again — it is safe to repeat.', ui.ButtonSet.OK); else throw err; }
}
function goAdvanceRegister() {
  const st = SpreadsheetApp.getActive().getSheetByName('Settings');
  st.activate(); st.setActiveRange(st.getRange(APP.ADV_FIRST, 1));
}
function goEmployeeMaster() {
  const st = SpreadsheetApp.getActive().getSheetByName('Settings');
  st.activate(); st.setActiveRange(st.getRange(APP.MASTER_FIRST, 15));
}

/* ========================= v2.3 EMPLOYEE CAPACITY ========================= */
/** Menu: grow (or re-lay out) the file for more employees. Keeps every employee, attendance entry, holiday,
 *  setting, shift change, increment, advance and locked month. Runs in stages and continues automatically. */
function changeCapacity() {
  const ui = SpreadsheetApp.getUi(), ss = SpreadsheetApp.getActive(), props = PropertiesService.getDocumentProperties();
  if (props.getProperty('CAP_STEP') !== null) { ui.alert('A capacity change is already running. Wait for "Capacity updated" (progress shows in pop-up messages).'); return; }
  const st = ss.getSheetByName('Settings'); if (!st) { ui.alert('Run setupPayrollSystem first.'); return; }
  const used = masterRowsUsed_(st, APP);
  const res = ui.prompt('Change employee capacity', `Current capacity: ${APP.ROWS} employees. Rows with an employee in the Employee Master: ${used}.\n\n` +
    'Enter the new capacity (e.g. 100, 150, 200 — maximum 300).\nAll data is kept. Takes about 3–6 minutes and continues automatically.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const cap = Math.round(Number(String(res.getResponseText()).trim()));
  if (!(cap >= 10 && cap <= 300)) { ui.alert('Enter a number between 10 and 300.'); return; }
  if (cap < used) { ui.alert(`You have ${used} employee rows — choose a capacity of at least ${used}.`); return; }
  if (cap === APP.ROWS && !APP.LEGACY) { ui.alert('Capacity is already ' + cap + '.'); return; }
  props.setProperties({ CAP_OLD: String(APP.ROWS), CAP_OLD_LEGACY: APP.LEGACY ? '1' : '0', CAP_NEW: String(cap), CAP_STEP: '0' });
  runCapSteps_();
}
function continueCapacity() { runCapSteps_(); }
function masterRowsUsed_(st, L) {
  const last = Math.min(st.getMaxRows(), L.CHG_TITLE - 1);
  const v = st.getRange(APP.MASTER_FIRST, 1, last - APP.MASTER_FIRST + 1, 1).getValues();
  let n = 0; v.forEach((r, i) => { if (String(r[0]).trim() !== '') n = i + 1; });
  return n;
}
function capSteps_() {
  const ss = SpreadsheetApp.getActive(), p = PropertiesService.getDocumentProperties();
  const old = layoutFor_(Number(p.getProperty('CAP_OLD')), p.getProperty('CAP_OLD_LEGACY') === '1');
  const steps = [['Settings & Employee Master', () => relayoutSettings_(old, Number(p.getProperty('CAP_NEW')))]];
  MONTHS.forEach((m, i) => steps.push(['Month tab ' + m, () => relayoutMonth_(ss.getSheetByName(m), i + 1, old)]));
  steps.push(['Salary Slip', () => recreateSheet_('Salary Slip', buildSalarySlip_, ['C2', 'F2'])]);
  steps.push(['Payroll Register', () => recreateSheet_(REG, buildRegister_, ['C2'])]);
  steps.push(['Dashboard', () => recreateSheet_('Dashboard', buildDashboard_, ['N15'])]);
  steps.push(['Recalculating payroll', () => { defineNamedRanges_(ss); syncEmployees(); markOffDays(); recalcAll_(); renderRegister_(); }]);
  steps.push(['Colours, formats & protection', () => { refreshAttendanceFormatting(); applyCurrencyFormats(); applyProtections_(ss); installTriggers_(); }]);
  return steps;
}
function runCapSteps_() {
  const t0 = Date.now(), p = PropertiesService.getDocumentProperties(), ui = safeUi_();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { ScriptApp.newTrigger('continueCapacity').timeBased().after(60 * 1000).create(); return; }
  try {
    ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'continueCapacity') ScriptApp.deleteTrigger(t); });
    const steps = capSteps_();
    let i = Number(p.getProperty('CAP_STEP') || 0), done = 0;
    while (i < steps.length) {
      if (done > 0 && Date.now() - t0 > APP.TIME_BUDGET_MS) {
        ScriptApp.newTrigger('continueCapacity').timeBased().after(20 * 1000).create();
        toast_(`Capacity change paused at step ${i + 1} of ${steps.length} — continues automatically in about 20 seconds. Keep the file open; do not edit.`, 'Payroll', 30);
        return;
      }
      toast_(`Changing capacity — step ${i + 1} of ${steps.length}: ${steps[i][0]}… (please wait, do not edit)`, 'Payroll', 30);
      try { steps[i][1](); }
      catch (err) { toast_('Capacity change stopped at "' + steps[i][0] + '": ' + err.message + ' — run Payroll System ▸ Change employee capacity again to continue.', 'Payroll', 30);
        p.deleteProperty('CAP_STEP_LOCK'); throw err; }
      SpreadsheetApp.flush(); i++; done++; p.setProperty('CAP_STEP', String(i));
    }
    ['CAP_STEP', 'CAP_OLD', 'CAP_OLD_LEGACY', 'CAP_NEW'].forEach(k => p.deleteProperty(k));
    toast_('Capacity updated ✔ — you can now add up to ' + APP.ROWS + ' employees.', 'Payroll', 20);
    if (ui) try { ui.alert('Capacity updated ✔', `The file now holds up to ${APP.ROWS} employees.\nAll employees, attendance, holidays, salary changes and advances were kept.\n\nReload the browser tab once, then run Payroll System ▸ Run Self-Test.`, ui.ButtonSet.OK); } catch (e) { }
  } finally { lock.releaseLock(); }
}
/** Settings: snapshot every input at the OLD positions, rebuild the tab at the NEW capacity, write the inputs back. */
function relayoutSettings_(O, newCap) {
  const ss = SpreadsheetApp.getActive(), st = ss.getSheetByName('Settings');
  const read = (r, c, nr, nc) => { if (r + nr - 1 > st.getMaxRows() || nr < 1) return { v: [], f: [] }; const g = st.getRange(r, c, nr, nc); return { v: g.getValues(), f: g.getFormulas() }; };
  const snap = {
    comp: read(4, 3, 17, 1), shifts: read(5, 6, 4, 3), offs: read(12, 7, 7, 4), stat: read(5, 12, 5, 3), heads: read(23, 18, 1, 7),
    hol: read(5, APP.HOL_COL, APP.HOL_ROWS, 3),
    master: read(APP.MASTER_FIRST, 1, O.CHG_TITLE - APP.MASTER_FIRST, APP.EXIT_COL),
    chgA: read(O.CHG_FIRST, 1, APP.CHG_ROWS, 4), chgH: read(O.CHG_FIRST, 8, APP.CHG_ROWS, 4),
    adv: read(O.ADV_FIRST, 1, APP.ADV_ROWS, 26), advOn: O.ADV_LAST <= st.getMaxRows() && String(st.getRange(O.ADV_TITLE, 1).getValue()).indexOf('I.') === 0,
    banner: st.getRange('A22').getValue()
  };
  // switch the whole script to the new layout
  PropertiesService.getDocumentProperties().setProperty('CAPACITY', String(newCap));
  applyLayout_(newCap, false);
  const all = st.getRange(1, 1, st.getMaxRows(), st.getMaxColumns());
  all.breakApart(); all.clearDataValidations(); all.clearNote();
  st.setConditionalFormatRules([]);
  st.setRowHeights(1, st.getMaxRows(), 21);
  buildSettings_(st);
  const put = (r, c, rows) => { if (rows.length) st.getRange(r, c, rows.length, rows[0].length).setValues(rows); };
  // company settings: values only (formula defaults stay)
  snap.comp.v.forEach((row, i) => { const r = 4 + i; if (r === 14 || snap.comp.f[i][0]) return; st.getRange(r, 3).setValue(row[0]); });
  put(5, 6, snap.shifts.v); put(12, 7, snap.offs.v); put(5, 12, snap.stat.v); put(23, 18, snap.heads.v); put(5, APP.HOL_COL, snap.hol.v);
  // employee master (inputs only; a gross typed without heads moves into Basic)
  const rows = snap.master.v.filter(r => r.some((x, k) => [6, 8, 15, 16, 25].indexOf(k) < 0 && String(x).trim() !== ''));
  if (rows.length > APP.ROWS) throw new Error(`${rows.length} employee rows do not fit in ${APP.ROWS}.`);
  rows.forEach(r => { if (r.slice(17, 24).every(x => x === '' || x === null) && Number(r[6]) > 0) r[17] = Number(r[6]); });
  const F = APP.MASTER_FIRST;
  if (rows.length) {
    put(F, 1, rows.map(r => r.slice(0, 6))); put(F, 8, rows.map(r => [r[7]])); put(F, 10, rows.map(r => r.slice(9, 15)));
    put(F, 18, rows.map(r => r.slice(17, 25))); put(F, APP.EXIT_COL, rows.map(r => [r[26]]));
  }
  put(APP.CHG_FIRST, 1, snap.chgA.v); put(APP.CHG_FIRST, 8, snap.chgH.v);
  if (snap.advOn && snap.adv.v.length) {
    put(APP.ADV_FIRST, 1, snap.adv.v.map(r => [r[0]])); put(APP.ADV_FIRST, 3, snap.adv.v.map(r => r.slice(2, 10))); put(APP.ADV_FIRST, 14, snap.adv.v.map(r => r.slice(13, 25)));
  }
  if (!/SAMPLE/i.test(String(snap.banner))) st.getRange('A22').setValue(snap.banner).setBackground(APP.NAVY2);
  defineNamedRanges_(ss);
}
/** Month tab: keep its employee rows (attendance, other earnings, frozen figures of a locked month), rebuild at the new size. */
function relayoutMonth_(sh, m, O) {
  const cols = Math.min(sh.getMaxColumns(), OUT_LAST);
  const data = sh.getRange(APP.M_FIRST, 1, O.ROWS, cols).getValues();
  const lockTxt = String(sh.getRange('AG1').getValue());
  sh.showRows(1, sh.getMaxRows());
  const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()); all.breakApart(); all.clearDataValidations();
  sh.setConditionalFormatRules([]);
  buildMonth_(sh, m);
  sh.getRange(APP.M_FIRST, 1, data.length, cols).setValues(data);
  if (lockTxt.indexOf('🔒') === 0) sh.getRange('AG1:AL1').merge().setValue(lockTxt).setBackground(APP.RED).setFontColor(APP.WHITE).setFontWeight('bold').setHorizontalAlignment('center');
}
/** Delete and rebuild a tab that holds no typed data (keeps its selector cells). */
function recreateSheet_(name, builder, keepCells) {
  const ss = SpreadsheetApp.getActive(), old = ss.getSheetByName(name);
  const keep = {}, idx = old ? old.getIndex() : ss.getSheets().length;
  if (old) { keepCells.forEach(a => { keep[a] = old.getRange(a).getValue(); }); ss.deleteSheet(old); }
  const sh = ss.insertSheet(name, Math.max(0, idx - 1));
  builder(sh);
  keepCells.forEach(a => { if (keep[a] !== '' && keep[a] !== undefined) sh.getRange(a).setValue(keep[a]); });
}

/* ============================== SAMPLE DATA ================================ */
function writeSampleData_(ss) {
  const st = ss.getSheetByName('Settings');
  const tz = ss.getSpreadsheetTimeZone();
  const year = Number(st.getRange('C8').getValue());
  // Sample holidays (verify dates for your state / company every year)
  const hol = [[serial_(year, 1, 26), 'Republic Day'], [serial_(year, 3, 4), 'Holi (sample — verify)'],
    [serial_(year, 8, 15), 'Independence Day'], [serial_(year, 9, 14), 'Ganesh Chaturthi (sample — verify)'],
    [serial_(year, 10, 2), 'Gandhi Jayanti'], [serial_(year, 10, 20), 'Dussehra (sample — verify)'],
    [serial_(year, 11, 9), 'Diwali holiday (sample — verify)'], [serial_(year, 12, 25), 'Christmas']];
  st.getRange(5, APP.HOL_COL, hol.length, 2).setValues(hol);

  // 5 fictional employees  (A..F, H, then J..O; G and I are formulas)
  //  EPF: Applicable, Employee share, Employer share (blank = defaults in Settings)
  const emps = [
    ['EMP001', 'Employee 1', 'Accounts', 'Accountant', 'Day', serial_(year - 2, 4, 1), 30000, MDEF, 'Yes', 1800, 1800, 0, 'employee1@example.com', 'Active'],
    ['EMP002', 'Employee 2', 'Operations', 'Executive', 'Day', serial_(year - 1, 6, 15), 22000, MDEF, 'No', '', '', 500, 'employee2@example.com', 'Active'],
    ['EMP003', 'Employee 3', 'Support', 'Night Shift Analyst', 'Night', serial_(year, 1, 10), 35000, MDEF, 'Yes', '', '', 0, 'employee3@example.com', 'Active'],
    ['EMP004', 'Employee 4', 'Support', 'Associate', 'Night', serial_(year - 1, 2, 1), 25000, MWD, 'Yes', 1250, 1250, 0, 'employee4@example.com', 'Active'],
    ['EMP005', 'Employee 5', 'Accounts', 'Trainee', 'Day (5-day)', serial_(year, 9, 15), 18000, MDEF, 'No', '', '', 0, 'employee5@example.com', 'Active']
  ];
  const F = APP.MASTER_FIRST;
  st.getRange(F, 1, emps.length, 6).setValues(emps.map(e => e.slice(0, 6)));     // G (gross) is a formula = sum of heads
  st.getRange(F, 8, emps.length, 1).setValues(emps.map(e => [e[7]]));
  // salary structure: Basic, DA, Special, Laptop, CEA, Others, Medical, ESI Employer
  const heads = {
    EMP001: [15000, 10000, 4000, 1000, 0, 0, 0, 0], EMP002: [11000, 7000, 3000, 1000, 0, 0, 0, 0],
    EMP003: [17500, 10000, 6500, 1000, 0, 0, 0, 0], EMP004: [15000, 10000, 0, 0, 0, 0, 0, 0],
    EMP005: [9000, 6000, 1000, 0, 0, 1000, 1000, 0] };
  st.getRange(F, 18, emps.length, 8).setValues(emps.map(e => heads[e[0]]));
  st.getRange(F, 10, emps.length, 6).setValues(emps.map(e => e.slice(8)));
  // sample mid-month changes in the current month: EMP002 moves Day → Night, EMP003 gets +₹2,000
  const cmS = Number(Utilities.formatDate(new Date(), tz, 'M'));
  st.getRange(APP.CHG_FIRST, 1, 1, 4).setValues([['EMP002', serial_(year, cmS, 16), 'Night', 'Moved to night shift (sample)']]);
  st.getRange(APP.CHG_FIRST, 8, 1, 4).setValues([['EMP003', serial_(year, cmS, 16), 2000, 'Annual increment (SAMPLE)']]);
  // sample salary advance: EMP002 ₹6,000 recovered ₹2,000 a month from the current month (B is a formula — skipped)
  st.getRange(APP.ADV_FIRST, 1).setValue('EMP002');
  st.getRange(APP.ADV_FIRST, 3, 1, 8).setValues([[serial_(year, cmS, 5), 'Salary Advance', 6000, 'Monthly', 2000, serial_(year, cmS, 1), 0, 'SAMPLE — 3 monthly instalments']]);
}

/**
 * Writes one month's sample employees + attendance + Weekly Off / Holiday labels as plain values.
 * Everything is computed in script (no reads of month sheets), so it is fast during setup.
 */
function writeSampleMonth_(ss, mi) {
  const ctx = readContext_(ss);
  const tz = ss.getSpreadsheetTimeZone();
  const year = ctx.year, todayS = ctx.today, status = ctx.statuses;
  const order = ['EMP001', 'EMP002', 'EMP003', 'EMP004', 'EMP005'];
  const pick = (e, s) => {
    let h = (Math.imul(s, 2654435761) ^ Math.imul(e + 1, 40503)) >>> 0; h = h % 100;
    if (h < 3) return status[4]; if (h < 7) return status[3]; if (h < 12) return status[2]; if (h < 20) return status[1]; return status[0];
  };
  const force = {};   // guaranteed variety for EMP001 in the current month + one MISSING day for EMP004
  const cm = Number(Utilities.formatDate(new Date(), tz, 'M'));
  force['EMP001|' + serial_(year, cm, 3)] = status[1];
  force['EMP001|' + serial_(year, cm, 10)] = status[2];
  force['EMP001|' + serial_(year, cm, 16)] = status[3];
  force['EMP001|' + serial_(year, cm, 22)] = status[4];
  force['EMP004|' + (todayS - 1)] = '';                 // yesterday left blank → MISSING demo
  // sandwich demos in the previous month (night shift, Sat+Sun off): EMP003 Absent Fri + Mon, EMP004 Leave Fri + Mon
  if (cm > 1) {
    const pm = cm - 1, fr = [];
    for (let d = 1; d + 3 <= daysIn_(year, pm); d++) {
      const s0 = serial_(year, pm, d), k = shiftIdx_(ctx, 'Night');
      if (new Date(Date.UTC(1899, 11, 30) + s0 * 86400000).getUTCDay() === 5 && dayType_(ctx, s0, k) === 'Work' &&
          dayType_(ctx, s0 + 1, k) === 'Off' && dayType_(ctx, s0 + 2, k) === 'Off' && dayType_(ctx, s0 + 3, k) === 'Work') fr.push(s0);
    }
    if (fr[0]) { force['EMP003|' + fr[0]] = status[4]; force['EMP003|' + (fr[0] + 3)] = status[4]; }
    if (fr[1]) { force['EMP004|' + fr[1]] = status[3]; force['EMP004|' + (fr[1] + 3)] = status[3]; }
  }
  const m = mi + 1, dim = daysIn_(year, m), monthEnd = serial_(year, m, dim);
  // same membership rule as syncEmployees: active and joined on/before month end, in Employee Master order
  const ids = Object.keys(ctx.master).filter(id => ctx.master[id].status === 'Active' && (!ctx.master[id].join || ctx.master[id].join <= monthEnd));
  if (!ids.length) return;
  const grid = ids.map(id => {
    const e = ctx.master[id], ei = order.indexOf(id), row = [];
    for (let d = 1; d <= 31; d++) {
      let v = '';
      if (d <= dim) {
        const s = serial_(year, m, d);
        const t = dayType_(ctx, s, shiftIdx_(ctx, shiftOn_(ctx, id, s)));
        const last = (ei <= 2) ? todayS : todayS - 1;   // EMP004/005 not yet marked today
        if (s >= e.join) {
          if (t === 'Work' && s <= last && ei >= 0) { v = pick(ei, s); const k = id + '|' + s; if (k in force) v = force[k]; }
          else if (t === 'Off') v = OFF_LABEL;
          else if (t === 'Hol') v = HOL_LABEL;
        }
      }
      row.push(v);
    }
    return row;
  });
  const sh = ss.getSheetByName(MONTHS[mi]);
  sh.getRange(APP.M_FIRST, 1, ids.length, 1).setValues(ids.map(x => [x]));
  sh.getRange(APP.M_FIRST, 8, ids.length, 31).setValues(grid);
}

/** Removes the 5 sample employees and all attendance (keeps settings, holidays & formulas). */
function clearSampleData() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Clear sample data?', 'This removes ALL employees in the Employee Master and ALL attendance in every month sheet. Settings, holidays and formulas are kept.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActive();
  const st = ss.getSheetByName('Settings');
  st.getRange(APP.MASTER_FIRST, 1, APP.ROWS, 6).clearContent();                    // A..F (G is a formula)
  st.getRange(APP.MASTER_FIRST, 8, APP.ROWS, 1).clearContent();                    // H
  st.getRange(APP.MASTER_FIRST, 10, APP.ROWS, 6).clearContent();                   // J..O
  st.getRange(APP.MASTER_FIRST, 18, APP.ROWS, 8).clearContent();                   // R..Y salary heads & ESI
  st.getRange(APP.MASTER_FIRST, APP.EXIT_COL, APP.ROWS, 1).clearContent();          // AA last working day
  if (st.getMaxRows() >= APP.ADV_LAST) {                                            // I. Advance & Loan Register
    st.getRange(APP.ADV_FIRST, 1, APP.ADV_ROWS, 1).clearContent();
    st.getRange(APP.ADV_FIRST, 3, APP.ADV_ROWS, 8).clearContent();
    st.getRange(APP.ADV_FIRST, 14, APP.ADV_ROWS, 12).clearContent();
  }
  st.getRange(APP.CHG_FIRST, 1, APP.CHG_ROWS, 4).clearContent();
  st.getRange(APP.CHG_FIRST, 8, APP.CHG_ROWS, 4).clearContent();
  st.getRange('A22').setValue('EMPLOYEE MASTER — add your employees below (one row each). Only Active employees appear in month sheets & the Salary Slip.')
    .setBackground(APP.NAVY2);
  MONTHS.forEach(m => {
    const sh = ss.getSheetByName(m);
    if (isLocked_(sh)) sh.getRange('AG1:AL1').breakApart().clearContent().setBackground(APP.NAVY);
    sh.showRows(APP.M_FIRST, APP.ROWS);
    sh.getRange(APP.M_FIRST, 1, APP.ROWS, 1).clearContent();
    sh.getRange(APP.M_FIRST, 8, APP.ROWS, 31).clearContent();
    sh.getRange(APP.M_FIRST, colN_('AY'), APP.ROWS, 1).clearContent();
  });
  ss.getSheetByName('Salary Slip').getRange('C2').clearContent();
  recalcAll_(); renderRegister_();
  toast_('Sample data cleared. Add employees in Settings ▸ Employee Master.', 'Done', 8);
}

/* =========================== CURRENCY & PROTECTION ========================= */
function applyCurrencyFormats() {
  const ss = SpreadsheetApp.getActive();
  const st = ss.getSheetByName('Settings');
  const sym = String(st.getRange('C9').getValue() || '₹').replace(/"/g, '');
  const f0 = `"${sym}"#,##0;-"${sym}"#,##0`, f2 = `"${sym}"#,##0.00;-"${sym}"#,##0.00`;
  const MF = APP.MASTER_FIRST, ML = APP.MASTER_LAST, F = APP.M_FIRST, L = APP.M_LAST;
  st.getRangeList([`G${MF}:G${ML}`, `K${MF}:M${ML}`, `Q${MF}:Q${ML}`, `R${MF}:Z${ML}`, 'C12', 'C16', `J${APP.CHG_FIRST}:J${APP.CHG_LAST}`, `L${APP.CHG_FIRST}:L${APP.CHG_LAST}`]).setNumberFormat(f0);
  st.getRange(`I${MF}:I${ML}`).setNumberFormat(f2);
  if (st.getMaxRows() >= APP.ADV_LAST) { const A1 = APP.ADV_FIRST, A2 = APP.ADV_LAST;
    st.getRangeList([`E${A1}:E${A2}`, `G${A1}:G${A2}`, `I${A1}:I${A2}`, `K${A1}:L${A2}`, `N${A1}:Y${A2}`]).setNumberFormat(f0); }
  MONTHS.forEach(m => {
    const sh = ss.getSheetByName(m); if (!sh) return;
    sh.getRangeList([`F${F}:F${L}`, `AX${F}:BE${L}`, 'Q4:AF4', `EF${F}:EF${L}`]).setNumberFormat(f0);
    if (sh.getMaxColumns() >= OUT_LAST) sh.getRange(F, XC.ADV, APP.ROWS, 1).setNumberFormat(f0);
    sh.getRange(`G${F}:G${L}`).setNumberFormat(f2);
  });
  const d = ss.getSheetByName('Dashboard');
  if (d) d.getRangeList(['L21:Q33', 'D17:O17', 'L13:O13', `Q76:U${75 + APP.ROWS}`]).setNumberFormat(f0);
  const s = ss.getSheetByName('Salary Slip');
  if (s) s.getRangeList(['C19:C28', 'F19:F28', 'F30:F32']).setNumberFormat(f0);
  const g = ss.getSheetByName(REG);
  if (g) g.getRangeList([`E8:E${8 + APP.ROWS}`, `P8:Z${8 + APP.ROWS}`, 'C5:U5']).setNumberFormat(f0);
}

function applyProtections_(ss) {
  // Warning-only protection: editors see a warning before overwriting a formula.
  ss.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => { if (p.getDescription().indexOf(APP.TAG) === 0) p.remove(); });
  ss.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => { if (p.getDescription().indexOf(APP.TAG) === 0) p.remove(); });
  const F = APP.M_FIRST, L = APP.M_LAST, MF = APP.MASTER_FIRST, ML = APP.MASTER_LAST;
  MONTHS.forEach(m => {
    const sh = ss.getSheetByName(m);
    const p = sh.protect().setDescription(APP.TAG + ': formulas on ' + m);
    p.setUnprotectedRanges([sh.getRange(`H${F}:AL${L}`), sh.getRange(`AY${F}:AY${L}`)]);
    p.setWarningOnly(true);
  });
  const dsh = ss.getSheetByName('Dashboard');
  const pd = dsh.protect().setDescription(APP.TAG + ': Dashboard'); pd.setUnprotectedRanges([dsh.getRange('N15:O15')]); pd.setWarningOnly(true);
  const rsh = ss.getSheetByName(REG);
  const pr = rsh.protect().setDescription(APP.TAG + ': Payroll Register'); pr.setUnprotectedRanges([rsh.getRange('C2:D2')]); pr.setWarningOnly(true);
  const ssh = ss.getSheetByName('Salary Slip');
  const ps = ssh.protect().setDescription(APP.TAG + ': Salary Slip'); ps.setUnprotectedRanges([ssh.getRange('C2:D2'), ssh.getRange('F2')]); ps.setWarningOnly(true);
  const st = ss.getSheetByName('Settings');
  [`G${MF}:G${ML}`, `I${MF}:I${ML}`, `P${MF}:Q${ML}`, `Z${MF}:Z${ML}`, 'I5:J8', 'G11:J11', `AF5:AF${ML}`, 'AH5:AI16', 'C14:D14',
   `E${APP.CHG_FIRST}:E${APP.CHG_LAST}`, `L${APP.CHG_FIRST}:L${APP.CHG_LAST}`]
   .concat(st.getMaxRows() >= APP.ADV_LAST ? [`B${APP.ADV_FIRST}:B${APP.ADV_LAST}`, `K${APP.ADV_FIRST}:Z${APP.ADV_LAST}`] : []).forEach(a1 =>
    st.getRange(a1).protect().setDescription(APP.TAG + ': auto ' + a1).setWarningOnly(true));
}

/* ================================ SELF-TEST ================================ */
/**
 * Re-calculates every employee/month independently in Apps Script and compares
 * with the spreadsheet values, then runs change-scenarios (leap year, WFH %,
 * EPF, salary slip switching). Everything changed during the test is restored.
 */
function runSelfTest() {
  const ss = SpreadsheetApp.getActive();
  const st = ss.getSheetByName('Settings');
  const slip = ss.getSheetByName('Salary Slip');
  const out = [];
  const ok = (cond, label) => out.push((cond ? '✅ ' : '❌ ') + label);
  SpreadsheetApp.flush();

  // 1. full recomputation of all month sheets
  const ctx = readContext_(ss);
  let rows = 0, bad = [];
  const allV = MONTHS.map(m => ss.getSheetByName(m).getRange(APP.M_FIRST, 1, APP.ROWS, OUT_LAST).getValues());
  MONTHS.forEach((m, i) => {
    const vals = allV[i];
    const res = verifyMonthValues_(ctx, i + 1, vals, allV);
    rows += res.rows; bad = bad.concat(res.errors.map(e => m + ': ' + e));
  });
  ok(bad.length === 0, `Payroll maths: ${rows} employee-month rows re-calculated independently` + (bad.length ? ' — ' + bad.slice(0, 5).join(' | ') : ''));
  const advErr = verifyAdvances_(ctx, allV);
  ok(advErr.length === 0, `Advance & loan recovery re-simulated month by month (${Object.keys(ctx.adv).length} employee(s) with advances)` + (advErr.length ? ' — ' + advErr.slice(0, 4).join(' | ') : ''));

  // 2. header dates / leap year
  const yCell = st.getRange('C8'); const y0 = yCell.getValue();
  const feb = ss.getSheetByName('February');
  let leap, leapDate, norm, normDate;
  try {
    yCell.setValue(2028); SpreadsheetApp.flush(); leap = feb.getRange('F2').getValue(); leapDate = feb.getRange('AJ7').getDisplayValue();
    yCell.setValue(2027); SpreadsheetApp.flush(); norm = feb.getRange('F2').getValue(); normDate = feb.getRange('AJ7').getDisplayValue();
  } finally { yCell.setValue(y0); SpreadsheetApp.flush(); }   // year is ALWAYS restored
  ok(leap === 29 && leapDate !== '', 'Leap-year February (2028) has 29 days');
  ok(norm === 28 && normDate === '', 'Normal February (2027) has 28 days; day 29 is blank/greyed');

  // 3. weekly off + holiday recognised
  const sep = ss.getSheetByName(MONTHS[8]);
  const hdr = sep.getRange('H7:AL12').getValues();
  let offOK = true, holOK = true;
  for (let d = 0; d < 31; d++) {
    const s = toSerial_(hdr[0][d], ctx.tz); if (!s) continue;
    for (let k = 0; k < 3; k++) { if (hdr[2 + k][d] !== dayType_(ctx, s, k)) { if (ctx.holidays[s]) holOK = false; else offOK = false; } }
  }
  ok(offOK, 'Weekly offs per shift match Settings (September)');
  ok(holOK, 'Holidays from the Holiday List are recognised (September)');

  // 4. WFH % change flows through
  // FIX v1.2: the original WFH % is ALWAYS restored, even if the test stops part-way
  const wfhCell = st.getRange('M6'); const w0 = wfhCell.getValue();
  let r2 = { errors: ['not run'] };
  try {
    wfhCell.setValue(0.6); recalcMonths_([8, 9, 10]); SpreadsheetApp.flush();
    const ctx2 = readContext_(ss);
    const sepV = sep.getRange(APP.M_FIRST, 1, APP.ROWS, OUT_LAST).getValues();
    r2 = verifyMonthValues_(ctx2, 9, sepV, allV.map((x, i) => i === 8 ? sepV : x));
  } finally {
    wfhCell.setValue(w0); recalcAll_(); SpreadsheetApp.flush();
  }
  ok(r2.errors.length === 0, 'Changing WFH % (test only, original restored) recalculates September correctly');

  // 5. EPF rules are part of the per-row check above
  ok(bad.length === 0, 'EPF: No → ₹0; Yes → employee share (+ employer share when borne from CTC); blank → defaults');

  // 6. salary slip switching employee + month
  const c0 = slip.getRange('C2').getValue(), f0 = slip.getRange('F2').getValue();
  const list = named_(ss, 'ActiveList').getValues().map(r => r[0]).filter(String);
  let slipOK = true, tested = 0;
  [MONTHS[8], MONTHS[7]].forEach(mn => {
    const msh = ss.getSheetByName(mn);
    const ids = msh.getRange(APP.M_FIRST, 1, APP.ROWS, 1).getValues().map(r => String(r[0]));
    list.forEach(label => {
      slip.getRange('C2').setValue(label); slip.getRange('F2').setValue(mn); SpreadsheetApp.flush();
      const id = String(label).split(' - ')[0]; const k = ids.indexOf(id);
      const status = slip.getRange('K6').getValue();
      if (k < 0 || status !== 'OK') return;
      const net = msh.getRange(APP.M_FIRST + k, colN_('BE')).getValue();
      if (Math.abs(Number(slip.getRange('F32').getValue()) - Number(net)) > 0.5) slipOK = false;
      if (Math.abs(Number(slip.getRange('C28').getValue()) - [19, 20, 21, 22, 23, 24, 25, 26, 27].reduce((a, r) => a + Number(slip.getRange('C' + r).getValue()), 0)) > 0.5) slipOK = false;
      tested++;
    });
  });
  slip.getRange('C2').setValue(c0); slip.getRange('F2').setValue(f0); SpreadsheetApp.flush();
  ok(slipOK && tested > 0, `Salary Slip follows employee + month selection (${tested} slips compared)`);

  // 6b. payroll register totals = month sheet
  const reg = ss.getSheetByName(REG), rc0 = reg.getRange('C2').getValue();
  reg.getRange('C2').setValue(MONTHS[8]); renderRegister_(); SpreadsheetApp.flush();
  const regOK = Math.abs(Number(reg.getRange('Z8').getValue()) - Number(sep.getRange('AC4').getValue())) < 0.5 &&
    Math.abs(Number(reg.getRange('R8').getValue()) - Number(sep.getRange('Q4').getValue())) < 0.5;
  reg.getRange('C2').setValue(rc0); renderRegister_(); SpreadsheetApp.flush();
  ok(regOK, 'Payroll Register (September) totals match the September sheet');

  // 7. no error values anywhere
  const errs = [];
  ss.getSheets().forEach(s => {
    const dv = s.getDataRange().getDisplayValues();
    dv.forEach((row, ri) => row.forEach((v, ci) => { if (/^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|ERROR!|NUM!)/.test(v)) errs.push(s.getName() + '!' + colL_(ci + 1) + (ri + 1)); }));
  });
  ok(errs.length === 0, 'No #N/A / #REF! / #VALUE! / #DIV/0! / #NAME? anywhere' + (errs.length ? ' — ' + errs.slice(0, 8).join(', ') : ''));

  const passed = out.filter(x => x.indexOf('✅') === 0).length;
  const text = `Self-test: ${passed} of ${out.length} checks passed\n\n` + out.join('\n');
  try { SpreadsheetApp.getUi().alert('Payroll System — Self-Test', text, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { Logger.log(text); }
  return out;
}

/** Reads all settings into a plain object (used by the engine, self-test and sample data). */
function readContext_(ss) {
  const st = ss.getSheetByName('Settings');
  const tz = ss.getSpreadsheetTimeZone();
  const lastRow = Math.min(st.getMaxRows(), APP.ADV_LAST);          // older files (before the v2 upgrade) are shorter
  const v = st.getRange('A1:AI' + lastRow).getValues();
  const cell = a1 => { const c = colN_(a1.replace(/\d+/, '')) - 1, r = Number(a1.replace(/\D+/, '')) - 1; return v[r][c]; };
  const holidays = {};
  for (let r = 5; r <= 4 + APP.HOL_ROWS; r++) {
    const s = toSerial_(v[r - 1][APP.HOL_COL - 1], tz);
    if (s) (holidays[s] = holidays[s] || []).push(String(v[r - 1][APP.HOL_COL + 1]).trim());
  }
  const chg = {}, inc = {};
  for (let r = APP.CHG_FIRST; r <= APP.CHG_LAST; r++) {
    const x = v[r - 1];
    const cid = String(x[0]).trim(), cd = toSerial_(x[1], tz);
    if (cid && cd && String(x[2]).trim()) (chg[cid] = chg[cid] || []).push({ date: cd, shift: String(x[2]).trim() });
    const iid = String(x[7]).trim(), idt = toSerial_(x[8], tz);
    if (iid && idt && x[9] !== '' && !isNaN(Number(x[9]))) (inc[iid] = inc[iid] || []).push({ date: idt, amt: Number(x[9]) });
  }
  const master = {};
  for (let r = APP.MASTER_FIRST; r <= APP.MASTER_LAST; r++) {
    const x = v[r - 1]; const id = String(x[0]).trim(); if (!id) continue;
    const heads = x.slice(17, 24).reduce((a, q) => a + (Number(q) || 0), 0);   // Monthly Gross = Basic + DA + allowances
    master[id] = { name: String(x[1]), dept: String(x[2]), desig: String(x[3]), shift: String(x[4]).trim(), join: toSerial_(x[5], tz) || 0, salary: heads,
      method: String(x[7]), epsApp: String(x[9]), epsAmt: Number(x[10]) || 0, erAmt: Number(x[11]) || 0,
      other: Number(x[12]) || 0, status: String(x[14]).trim() || 'Active', esi: Number(x[24]) || 0 };
    master[id].exit = EXITED.indexOf(master[id].status) >= 0 ? (toSerial_(x[APP.EXIT_COL - 1], tz) || 0) : 0;   // last working day
  }
  // v2: Advance & Loan Register (Settings ▸ I). rec[0..11] = amount recovered in Jan..Dec (shared with advRows → written back)
  const adv = {}, advRows = [];
  const hasAdv = lastRow >= APP.ADV_LAST && String(v[APP.ADV_TITLE - 1][0]).indexOf('I.') === 0;
  if (hasAdv) for (let r = APP.ADV_FIRST; r <= APP.ADV_LAST; r++) {
    const x = v[r - 1], idx = r - APP.ADV_FIRST;
    const aid = String(x[0]).trim(), amt = Number(x[4]) || 0, dt = toSerial_(x[2], tz);
    if (!aid || !master[aid] || amt <= 0 || !dt) { advRows.push(new Array(12).fill('')); continue; }
    const rec = x.slice(13, 25).map(q => Number(q) || 0);
    advRows.push(rec);
    (adv[aid] = adv[aid] || []).push({ idx: idx, date: dt, amount: amt, already: Number(x[8]) || 0,
      mode: String(x[5]).trim() === 'One-time' ? 'One-time' : 'Monthly', inst: Number(x[6]) || 0,
      from: toSerial_(x[7], tz) || dt, rec: rec });
  }
  Object.keys(adv).forEach(k => adv[k].sort((a, b) => a.date - b.date || a.idx - b.idx));   // oldest advance is recovered first
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd').split('-').map(Number);
  return {
    tz: tz, year: Number(cell('C8')), defMethod: String(cell('C11')), defEPS: Number(cell('C12')) || 0,
    defEPSEr: Number(cell('C16')) || 0, erFromEmp: String(cell('C17')),
    sandwichOn: String(cell('C18')), sandwichBorder: String(cell('C19')),
    sandwich: [5, 6, 7, 8, 9].filter(r => String(v[r - 1][13]) === 'Yes').map(r => String(v[r - 1][11])),
    trackStart: toSerial_(cell('C13'), tz) || 0,
    statuses: [5, 6, 7, 8, 9].map(r => String(v[r - 1][11])), pct: [5, 6, 7, 8, 9].map(r => Number(v[r - 1][12]) || 0),
    shiftNames: [5, 6, 7, 8].map(r => String(v[r - 1][5])),
    off: [12, 13, 14, 15, 16, 17, 18].map(r => [6, 7, 8, 9].map(c => String(v[r - 1][c]) === 'Yes')),
    holidays: holidays, master: master, chg: chg, inc: inc, today: serial_(todayStr[0], todayStr[1], todayStr[2]),
    hasAdv: hasAdv, adv: adv, advRows: advRows, sym: String(cell('C9') || '₹'),
    advCap: (cell('C20') === '' || isNaN(Number(cell('C20')))) ? 1 : Math.min(1, Math.max(0, Number(cell('C20'))))
  };
}

/** Shift name in force for employee `id` on date serial s (base shift + latest change on/before s). */
function shiftOn_(ctx, id, s) {
  let best = null;
  (ctx.chg[id] || []).forEach(c => { if (c.date <= s && (!best || c.date > best.date)) best = c; });
  const e = ctx.master[id];                       // v2.1: an ID no longer in the Employee Master is skipped safely
  return best ? best.shift : (e ? e.shift : '');
}
/** Monthly salary in force for employee `id` on date serial s (base + increments effective on/before s). */
function rateOn_(ctx, id, s) {
  return (ctx.inc[id] || []).reduce((a, x) => a + (x.date <= s ? x.amt : 0), ctx.master[id] ? ctx.master[id].salary : 0);
}
function shiftIdx_(ctx, name) { return Math.max(0, ctx.shiftNames.indexOf(name)); }

/** Day type for a date serial and shift index: 'Work' | 'Off' | 'Hol' */
function dayType_(ctx, s, k) {
  const h = ctx.holidays[s];
  if (h && h.some(a => a === '' || a === 'All' || a === ctx.shiftNames[k])) return 'Hol';
  const wd = new Date(Date.UTC(1899, 11, 30) + s * 86400000).getUTCDay();   // 0=Sun
  return ctx.off[(wd + 6) % 7][k] ? 'Off' : 'Work';
}

/** Independent recomputation of one month sheet's values. vals = A..FO rows; all = the 12 months' vals (for sandwich across months). */
function verifyMonthValues_(ctx, m, vals, all) {
  all = all || [];
  if (!ctx._entries) {                                          // id → {serial: entry} for the whole year
    ctx._entries = {};
    all.forEach((mv, mi) => { if (!mv) return; mv.forEach(rw => { const id = String(rw[0]).trim(); if (!id) return;
      const o = ctx._entries[id] = ctx._entries[id] || {};
      for (let d = 1; d <= daysIn_(ctx.year, mi + 1); d++) o[serial_(ctx.year, mi + 1, d)] = String(rw[7 + d - 1]).trim(); }); });
  }
  const C = x => colN_(x) - 1;
  const days = daysIn_(ctx.year, m);
  const errors = []; let rows = 0;
  const near = (a, b, t) => Math.abs((Number(a) || 0) - (Number(b) || 0)) <= t;
  vals.forEach((row, i) => {
    const id = String(row[0]).trim(); if (!id) return;
    const e = ctx.master[id]; if (!e) return;
    rows++;
    const method = (e.method === '' || e.method === MDEF) ? ctx.defMethod : e.method;
    const is30 = method === M30;
    const cnt = [0, 0, 0, 0, 0]; let eligCal = 0, eligWD = 0, shiftWD = 0, missing = 0;
    let sumCal = 0, sumWD = 0, unpaidAmt = 0;
    const ex = e.exit || 0, el = s => s >= e.join && (!ex || s <= ex);
    const onPay = s => (e.status === 'Active' || (EXITED.indexOf(e.status) >= 0 && ex > 0)) && el(s);
    for (let d = 1; d <= days; d++) {
      const s = serial_(ctx.year, m, d);
      const t = dayType_(ctx, s, shiftIdx_(ctx, shiftOn_(ctx, id, s)));
      const rate = rateOn_(ctx, id, s);
      const entry = String(row[7 + d - 1]).trim();
      if (t === 'Work') shiftWD++;
      if (el(s)) { eligCal++; sumCal += rate; if (t === 'Work') { eligWD++; sumWD += rate; } }
      if (t === 'Work' && el(s)) { const si = ctx.statuses.indexOf(entry); if (si >= 0) { cnt[si]++; unpaidAmt += (1 - ctx.pct[si]) * rate; } }
      if (onPay(s) && ctx.statuses.indexOf(entry) < 0 && t === 'Work' && s >= ctx.trackStart && s < ctx.today) missing++;
    }
    // ---- sandwich (same rules as the engine) ----
    const y0 = serial_(ctx.year, 1, 1), y1 = serial_(ctx.year, 12, 31);
    const typ = t => dayType_(ctx, t, shiftIdx_(ctx, shiftOn_(ctx, id, t)));
    const ent = t => ((ctx._entries[id] || {})[t] || '');
    const prevW = t => { for (let x = t - 1; x >= y0; x--) if (typ(x) === 'Work') return ent(x); return ''; };
    const nextW = t => { for (let x = t + 1; x <= y1; x++) if (typ(x) === 'Work') return ent(x); return ''; };
    const trig = x => ctx.sandwich.indexOf(x) >= 0;
    let swc = 0, swm = 0, brw = 0, brm = 0;
    const markedPre = cnt.reduce((a, b) => a + b, 0);
    if (ctx.sandwichOn === 'Yes' && markedPre > 0) {
      for (let d = 1; d <= days; d++) {
        const s = serial_(ctx.year, m, d), t = typ(s), entry = String(row[7 + d - 1]).trim();
        if (!el(s)) continue;
        if ((t === 'Off' || t === 'Hol') && ctx.statuses.indexOf(entry) < 0 && trig(prevW(s)) && trig(nextW(s))) { swc++; swm += rateOn_(ctx, id, s); }
        if (ctx.sandwichBorder === 'Unpaid (LOP)' && t === 'Work' && trig(entry)) {
          const offN = d < days && ['Off', 'Hol'].indexOf(typ(s + 1)) >= 0, offP = d > 1 && ['Off', 'Hol'].indexOf(typ(s - 1)) >= 0;
          if ((offN && trig(nextW(s))) || (offP && trig(prevW(s)))) { const pc = ctx.pct[ctx.statuses.indexOf(entry)]; brw += pc; brm += pc * rateOn_(ctx, id, s); }
        }
      }
    }
    const endRate = rateOn_(ctx, id, serial_(ctx.year, m, days));
    const marked = cnt.reduce((a, b) => a + b, 0);
    const paidW = cnt.reduce((a, c, j) => a + c * ctx.pct[j], 0);
    const daily = is30 ? endRate / 30 : (shiftWD ? endRate / shiftWD : 0);
    const unpaid = Math.round((marked - paidW + swc + brw) * 100) / 100;
    const earned = marked === 0 ? 0 : Math.round(is30 ? sumCal / days : (shiftWD ? sumWD / shiftWD : 0));
    const lopAmt = unpaidAmt + swm + brm;
    const absence = marked === 0 ? 0 : Math.min(earned, Math.round(is30 ? lopAmt / 30 : (shiftWD ? lopAmt / shiftWD : 0)));
    const esi = (marked === 0 || ctx.erFromEmp !== 'Yes') ? 0 : e.esi;
    const epfEE = (marked === 0 || e.epsApp !== 'Yes') ? 0 : (e.epsAmt > 0 ? e.epsAmt : ctx.defEPS);
    const epfER = (marked === 0 || e.epsApp !== 'Yes' || ctx.erFromEmp !== 'Yes') ? 0 : (e.erAmt > 0 ? e.erAmt : ctx.defEPSEr);
    const eps = epfEE + epfER;
    const other = marked === 0 ? 0 : e.other;
    const gross = earned + (Number(row[C('AY')]) || 0);
    const netPre = gross - absence - eps - other - esi;
    const advV = Number(row[XC.ADV - 1]) || 0;                       // advance recovery (schedule checked in verifyAdvances_)
    const net = netPre - advV;
    if (advV < -0.004 || advV > Math.max(0, netPre) + 0.5) errors.push(`${id} Advance recovery ${advV} is more than the net salary ${netPre}`);
    if (!near(row[XC.NETPRE - 1], netPre, 0.5)) errors.push(`${id} Net before advance: sheet ${row[XC.NETPRE - 1]} vs expected ${Math.round(netPre)}`);
    if (String(row[C('E')]) !== shiftOn_(ctx, id, serial_(ctx.year, m, days))) errors.push(`${id} Shift: sheet ${row[C('E')]} vs expected ${shiftOn_(ctx, id, serial_(ctx.year, m, days))}`);
    if (Math.abs(Number(row[C('F')]) - endRate) > 0.01) errors.push(`${id} Salary: sheet ${row[C('F')]} vs expected ${endRate}`);
    const checks = [['Present', C('AM'), cnt[0], 0], ['WFH', C('AN'), cnt[1], 0], ['HalfDay', C('AO'), cnt[2], 0], ['Leave', C('AP'), cnt[3], 0],
      ['Absent', C('AQ'), cnt[4], 0], ['Missing', C('AR'), missing, 0], ['Daily', C('G'), daily, 0.01], ['Earned', C('AX'), earned, 0.5],
      ['Gross', C('AZ'), gross, 0.5], ['Absence', C('BA'), absence, 0.5], ['EPF total', C('BB'), eps, 0.01], ['EPF employee', C('BR'), epfEE, 0.01],
      ['EPF employer', C('BS'), epfER, 0.01], ['Other', C('BC'), other, 0.01], ['ESI', C('EF'), esi, 0.01],
      ['Sandwich days', C('EG'), swc, 0.001], ['Border days', C('EI'), brw, 0.001],
      ['Net', C('BE'), net, 0.5], ['Unpaid', C('AU'), unpaid, 0.01]];
    checks.forEach(ch => { if (!near(row[ch[1]], ch[2], ch[3])) errors.push(`${id} ${ch[0]}: sheet ${row[ch[1]]} vs expected ${Math.round(ch[2] * 100) / 100}`); });
  });
  return { rows: rows, errors: errors };
}

/**
 * Independent re-simulation of the Advance & Loan Register: starting from the register inputs only, replays
 * every month in order using each month's net-before-advance and compares with what the engine deducted.
 * allV = the 12 month sheets' values (A..FS). Returns a list of error strings.
 */
function verifyAdvances_(ctx, allV) {
  const errors = [];
  if (!ctx.hasAdv) return errors;
  const y = ctx.year;
  Object.keys(ctx.adv).forEach(id => {
    const e = ctx.master[id];
    const list = ctx.adv[id].map(a => ({ a: a, paid: 0 }));
    for (let m = 1; m <= 12; m++) {
      const me = serial_(y, m, daysIn_(y, m)), ms = serial_(y, m, 1);
      const row = allV[m - 1].find(r => String(r[0]).trim() === id);
      let sheetAdv = 0, cap = 0, marked = 0;
      if (row) { sheetAdv = Number(row[XC.ADV - 1]) || 0; cap = Math.round(Math.max(0, Number(row[XC.NETPRE - 1]) || 0) * ctx.advCap); marked = Number(row[colN_('BL') - 1]) || 0; }
      const exitM = EXITED.indexOf(e.status) >= 0 && e.exit >= ms && e.exit <= me;
      let expect = 0;
      list.forEach(o => {
        const a = o.a, remaining = Math.round((a.amount - a.already - o.paid) * 100) / 100;
        const fd = new Date(Date.UTC(1899, 11, 30) + a.from * 86400000);
        const started = fd.getUTCFullYear() * 12 + fd.getUTCMonth() <= y * 12 + m - 1;
        let due = 0;
        if (row && remaining > 0 && a.date <= me && marked > 0) due = exitM ? remaining : (started ? (a.mode === 'One-time' ? remaining : Math.min(a.inst > 0 ? a.inst : remaining, remaining)) : 0);
        const take = Math.min(due, cap); cap -= take; o.paid += take; expect += take;
        if (Math.abs((Number(ctx.advRows[a.idx][m - 1]) || 0) - take) > 0.5) errors.push(`${id} advance row ${APP.ADV_FIRST + a.idx} ${MONTHS[m - 1]}: register ${ctx.advRows[a.idx][m - 1] || 0} vs expected ${take}`);
      });
      if (Math.abs(sheetAdv - expect) > 0.5) errors.push(`${id} ${MONTHS[m - 1]}: advance deducted on sheet ${sheetAdv} vs expected ${expect}`);
    }
  });
  return errors;
}

/* ================================ HELPERS ================================== */
function serial_(y, m, d) { return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000); }
function daysIn_(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }
function toSerial_(v, tz) {
  if (v === '' || v === null || v === undefined) return 0;
  if (typeof v === 'number') return Math.floor(v);
  if (Object.prototype.toString.call(v) === '[object Date]') {
    const p = Utilities.formatDate(v, tz, 'yyyy-MM-dd').split('-').map(Number); return serial_(p[0], p[1], p[2]);
  }
  const t = Date.parse(String(v)); if (isNaN(t)) return 0;
  const dd = new Date(t); return serial_(dd.getFullYear(), dd.getMonth() + 1, dd.getDate());
}
function colL_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function colN_(l) { let n = 0; for (let i = 0; i < l.length; i++) n = n * 26 + (l.charCodeAt(i) - 64); return n; }
function esc_(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function safeUi_() { try { return SpreadsheetApp.getUi(); } catch (e) { return null; } }
function toast_(msg, title, sec) { try { SpreadsheetApp.getActive().toast(msg, title || 'Payroll', sec || 5); } catch (e) { } }
function ensureSize_(sh, rows, cols) {
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
}
function styleBanner_(rg, size) {
  rg.setBackground(APP.NAVY).setFontColor(APP.WHITE).setFontSize(size || 14).setFontWeight('bold')
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  rg.setBorder(null, null, true, null, null, null, APP.RED, SpreadsheetApp.BorderStyle.SOLID_THICK);
}
function banner_(rg, text, size) { rg.merge().setValue(text); styleBanner_(rg, size); }
function section_(rg, text) {
  if (rg.getNumColumns() > 1) rg.merge();
  rg.setValue(text).setBackground(APP.PALE).setFontColor(APP.NAVY).setFontWeight('bold').setFontSize(10)
    .setBorder(null, true, null, null, null, null, APP.RED, SpreadsheetApp.BorderStyle.SOLID_THICK);
}
function styleHeader_(rg) {
  rg.setBackground(APP.NAVY).setFontColor(APP.WHITE).setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
}
function header_(rg, labels) {
  const f = labels.map(x => String(x).charAt(0) === '=' ? x : null);
  rg.setValues([labels.map(x => String(x).charAt(0) === '=' ? '' : x)]);
  f.forEach((x, i) => { if (x) rg.getCell(1, i + 1).setFormula(x); });
  styleHeader_(rg);
}
function box_(rg, color) { rg.setBorder(true, true, true, true, null, null, color || APP.BORDER, SpreadsheetApp.BorderStyle.SOLID); }
function card_(sh, labA1, valA1, label, formula, fmt, color) {
  const l = sh.getRange(labA1).merge();
  if (String(label).charAt(0) === '=') l.setFormula(label); else l.setValue(label);
  l.setBackground(APP.PALE).setFontColor(APP.MUTED).setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center');
  const v = sh.getRange(valA1).merge().setFormula(formula);
  v.setFontSize(18).setFontWeight('bold').setFontColor(color || APP.NAVY).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(APP.WHITE);
  if (fmt && fmt !== 'money' && fmt !== '@') v.setNumberFormat(fmt);
  sh.getRange(labA1.split(':')[0] + ':' + valA1.split(':')[1]).setBorder(true, true, true, true, null, null, APP.BORDER, SpreadsheetApp.BorderStyle.SOLID);
}