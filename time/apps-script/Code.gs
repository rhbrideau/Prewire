/**
 * Tightline Time: Google Sheets backend.
 *
 * Paste this whole file into Extensions > Apps Script of a new Google Sheet,
 * then follow time/SETUP.md. The sheet holds every entry; the phone app talks
 * to this script.
 */

const TAB = { ENTRIES: 'Entries', EMPLOYEES: 'Employees', SETTINGS: 'Settings', PAYROLL: 'Payroll' };
const ENTRY_COLS = ['ID', 'Employee', 'Date', 'Start', 'End', 'Break (min)', 'Hours', 'Job', 'Notes', 'Created', 'Updated'];
const EMP_COLS = ['Name', 'PIN', 'Hourly Rate', 'Email', 'Active'];
const PAY_COLS = ['Week Start', 'Week End', 'Employee', 'Regular Hours', 'Overtime Hours', 'Total Hours', 'Hourly Rate', 'Gross Pay', 'Approved At'];
const DEFAULT_SETTINGS = [
  ['OwnerPin', '', 'Your PIN. Opens the owner view with pay amounts. Keep it private.'],
  ['OwnerEmail', '', 'Where the Monday payroll summary is sent.'],
  ['AppUrl', '', 'Link to the time app. Goes in the reminder emails.'],
  ['ReminderHour', '18', 'Hour of day (0 to 23) for the daily reminder. Run Time Tracker > Setup again after changing.'],
  ['WorkDays', 'Mon,Tue,Wed,Thu,Fri', 'Days the "log your hours" reminder goes out.'],
  ['OvertimeAfter', '40', 'Weekly hours before overtime starts.'],
  ['OvertimeMultiplier', '1.5', 'Overtime pay multiplier.'],
];
const MAX_SHIFT_MIN = 16 * 60;

/* ---------------- Sheet menu and one time setup ---------------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Time Tracker')
    .addItem('Setup / reinstall reminders', 'setup')
    .addItem('Send me a test email', 'testEmail')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const entries = ensureSheet_(ss, TAB.ENTRIES, ENTRY_COLS);
  entries.getRange('A:F').setNumberFormat('@');
  entries.getRange('G:G').setNumberFormat('0.00');
  entries.getRange('H:K').setNumberFormat('@');

  const emp = ensureSheet_(ss, TAB.EMPLOYEES, EMP_COLS);
  emp.getRange('A:B').setNumberFormat('@');
  emp.getRange('C:C').setNumberFormat('0.00');
  emp.getRange('D:E').setNumberFormat('@');
  if (emp.getLastRow() < 2) emp.appendRow(['Employee', randomPin_(), 20, '', 'Yes']);

  const set = ensureSheet_(ss, TAB.SETTINGS, ['Setting', 'Value', 'Notes']);
  set.getRange('A:C').setNumberFormat('@');
  const have = set.getDataRange().getDisplayValues().map(r => r[0]);
  DEFAULT_SETTINGS.forEach(row => { if (have.indexOf(row[0]) === -1) set.appendRow(row); });
  const s = getSettings_();
  if (!s.OwnerPin) setSetting_('OwnerPin', randomPin_());
  if (!s.OwnerEmail) {
    try { setSetting_('OwnerEmail', Session.getEffectiveUser().getEmail()); } catch (e) { /* fill in by hand */ }
  }

  const pay = ensureSheet_(ss, TAB.PAYROLL, PAY_COLS);
  pay.getRange('A:C').setNumberFormat('@');
  pay.getRange('D:H').setNumberFormat('0.00');
  pay.getRange('I:I').setNumberFormat('@');

  installTriggers_();

  const msg = 'Setup done.\n\nYour owner PIN: ' + getSettings_().OwnerPin +
    '\n\nNext: open the Employees tab, put in his name, email, and hourly rate. ' +
    'His PIN is already filled in there. Then deploy as a web app (see SETUP.md).';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

function installTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const f = t.getHandlerFunction();
    if (f === 'dailyReminder' || f === 'weeklySummary') ScriptApp.deleteTrigger(t);
  });
  const hour = clampInt_(getSettings_().ReminderHour, 0, 23, 18);
  ScriptApp.newTrigger('dailyReminder').timeBased().everyDays(1).atHour(hour).create();
  ScriptApp.newTrigger('weeklySummary').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ---------------- Web app entry points ---------------- */

function doGet() {
  return ContentService.createTextOutput('Tightline Time is running. Paste this URL into the app.');
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const who = auth_(body.pin);
    let out;
    switch (body.action) {
      case 'auth': out = { role: who.role, name: who.name }; break;
      case 'week': out = getWeek_(who, body.weekStart); break;
      case 'save': out = saveEntry_(who, body.entry || {}, body.employee); break;
      case 'delete': out = deleteEntry_(who, body.id); break;
      case 'approve': out = approveWeek_(who, body.employee, body.weekStart, body.approved !== false); break;
      default: throw new Error('Unknown action');
    }
    out.ok = true;
    return json_(out);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- Auth ---------------- */

function auth_(pin) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pinfails') || 0);
  if (fails >= 10) throw new Error('Too many wrong PINs. Wait 15 minutes and try again.');
  pin = String(pin == null ? '' : pin).trim();
  if (!pin) throw new Error('Enter your PIN');

  const s = getSettings_();
  if (s.OwnerPin && pin === String(s.OwnerPin).trim()) return { role: 'owner', name: 'Owner' };
  const emp = getEmployees_().find(x => x.active && x.pin === pin);
  if (emp) return { role: 'employee', name: emp.name };

  cache.put('pinfails', String(fails + 1), 900);
  throw new Error('Wrong PIN');
}

/* ---------------- Data access ---------------- */

function sheet_(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }

function getSettings_() {
  const sh = sheet_(TAB.SETTINGS);
  const out = {};
  if (!sh) return out;
  sh.getDataRange().getDisplayValues().slice(1).forEach(r => { if (r[0]) out[r[0].trim()] = String(r[1]).trim(); });
  return out;
}

function setSetting_(key, value) {
  const sh = sheet_(TAB.SETTINGS);
  const rows = sh.getDataRange().getDisplayValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) { sh.getRange(i + 1, 2).setValue(String(value)); return; }
  }
  sh.appendRow([key, String(value), '']);
}

function getEmployees_() {
  const sh = sheet_(TAB.EMPLOYEES);
  if (!sh) return [];
  return sh.getDataRange().getDisplayValues().slice(1)
    .filter(r => String(r[0]).trim())
    .map(r => ({
      name: String(r[0]).trim(),
      pin: String(r[1]).trim(),
      rate: Number(String(r[2]).replace(/[^0-9.]/g, '')) || 0,
      email: String(r[3]).trim(),
      active: !/^(no|n|false|0)$/i.test(String(r[4]).trim()),
    }));
}

function readEntries_() {
  const sh = sheet_(TAB.ENTRIES);
  const rows = sh.getDataRange().getDisplayValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    out.push({
      row: i + 1, id: r[0], employee: r[1], date: r[2], start: r[3], end: r[4],
      breakMin: Number(r[5]) || 0, hours: Number(r[6]) || 0, job: unescapeCell_(r[7]), notes: unescapeCell_(r[8]),
    });
  }
  return out;
}

function publicEntry_(e) {
  return { id: e.id, employee: e.employee, date: e.date, start: e.start, end: e.end, breakMin: e.breakMin, hours: e.hours, job: e.job, notes: e.notes };
}

function approvedSet_() {
  const sh = sheet_(TAB.PAYROLL);
  const set = {};
  if (!sh) return set;
  sh.getDataRange().getDisplayValues().slice(1).forEach(r => { if (r[0] && r[8]) set[r[2] + '|' + r[0]] = true; });
  return set;
}

/* ---------------- Actions ---------------- */

function getWeek_(who, weekStart) {
  const ws = mondayOf_(weekStart || today_());
  const we = addDays_(ws, 6);
  const all = readEntries_();
  const inWeek = all.filter(e => e.date >= ws && e.date <= we).sort(byDateTime_);
  const approved = approvedSet_();
  const s = getSettings_();

  if (who.role === 'employee') {
    return {
      role: 'employee', name: who.name, weekStart: ws,
      entries: inWeek.filter(e => e.employee === who.name).map(publicEntry_),
      approved: !!approved[who.name + '|' + ws],
      recentJobs: recentJobs_(all.filter(e => e.employee === who.name)),
    };
  }

  const emps = getEmployees_();
  const names = emps.filter(x => x.active).map(x => x.name);
  inWeek.forEach(e => { if (names.indexOf(e.employee) === -1) names.push(e.employee); });
  return {
    role: 'owner', name: 'Owner', weekStart: ws,
    overtimeAfter: Number(s.OvertimeAfter) || 40,
    overtimeMultiplier: Number(s.OvertimeMultiplier) || 1.5,
    employees: names.map(n => {
      const rec = emps.find(x => x.name === n) || { rate: 0 };
      return {
        name: n, rate: rec.rate,
        approved: !!approved[n + '|' + ws],
        entries: inWeek.filter(e => e.employee === n).map(publicEntry_),
      };
    }),
    recentJobs: recentJobs_(all),
  };
}

function saveEntry_(who, raw, employeeForOwner) {
  const e = cleanEntry_(raw);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const approved = approvedSet_();
    const existing = readEntries_().find(x => x.id === e.id);
    let employee;
    if (who.role === 'owner') {
      employee = String(employeeForOwner || (existing && existing.employee) || '').trim();
      if (!employee) throw new Error('Pick an employee');
    } else {
      employee = who.name;
      if (existing && existing.employee !== who.name) throw new Error('That entry is not yours');
      if (approved[employee + '|' + mondayOf_(e.date)]) throw new Error('That week is approved and locked. Ask the owner to change it.');
      if (existing && approved[employee + '|' + mondayOf_(existing.date)]) throw new Error('That week is approved and locked. Ask the owner to change it.');
    }

    const now = stamp_();
    const row = [e.id, employee, e.date, e.start, e.end, String(e.breakMin), e.hours,
      escapeCell_(e.job), escapeCell_(e.notes), existing ? '' : now, now];
    const sh = sheet_(TAB.ENTRIES);
    if (existing) {
      row[9] = sh.getRange(existing.row, 10).getDisplayValue();
      sh.getRange(existing.row, 1, 1, row.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    const touched = [[employee, mondayOf_(e.date)]];
    if (existing) touched.push([existing.employee, mondayOf_(existing.date)]);
    refreshApproved_(touched);
    e.employee = employee;
    return { entry: e };
  } finally {
    lock.releaseLock();
  }
}

function deleteEntry_(who, id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const existing = readEntries_().find(x => x.id === String(id));
    if (!existing) return { deleted: false };
    if (who.role !== 'owner') {
      if (existing.employee !== who.name) throw new Error('That entry is not yours');
      if (approvedSet_()[who.name + '|' + mondayOf_(existing.date)]) throw new Error('That week is approved and locked. Ask the owner to change it.');
    }
    sheet_(TAB.ENTRIES).deleteRow(existing.row);
    refreshApproved_([[existing.employee, mondayOf_(existing.date)]]);
    return { deleted: true };
  } finally {
    lock.releaseLock();
  }
}

function approveWeek_(who, employee, weekStart, approve) {
  if (who.role !== 'owner') throw new Error('Only the owner can approve');
  employee = String(employee || '').trim();
  if (!employee) throw new Error('Pick an employee');
  const ws = mondayOf_(weekStart);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!approve) { removePayrollRow_(employee, ws); return { approved: false }; }
    return { approved: true, pay: writePayrollRow_(employee, ws) };
  } finally {
    lock.releaseLock();
  }
}

function findPayrollRow_(employee, ws) {
  const rows = sheet_(TAB.PAYROLL).getDataRange().getDisplayValues();
  for (let i = 1; i < rows.length; i++) if (rows[i][0] === ws && rows[i][2] === employee) return i + 1;
  return 0;
}

function removePayrollRow_(employee, ws) {
  const n = findPayrollRow_(employee, ws);
  if (n) sheet_(TAB.PAYROLL).deleteRow(n);
}

// Writes (or rewrites) the Payroll tab line for one employee and week. Caller holds the lock.
function writePayrollRow_(employee, ws) {
  const we = addDays_(ws, 6);
  const s = getSettings_();
  const rate = (getEmployees_().find(x => x.name === employee) || { rate: 0 }).rate;
  const mins = readEntries_().filter(e => e.employee === employee && e.date >= ws && e.date <= we)
    .reduce((t, e) => t + workMinutes_(e.start, e.end, e.breakMin), 0);
  const p = payFor_(mins, rate, Number(s.OvertimeAfter) || 40, Number(s.OvertimeMultiplier) || 1.5);
  const row = [ws, we, employee, p.regHours, p.otHours, p.totalHours, rate, p.gross, stamp_()];
  const sh = sheet_(TAB.PAYROLL);
  const n = findPayrollRow_(employee, ws);
  if (n) sh.getRange(n, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  return p;
}

// After the owner changes an entry inside an approved week, keep the Payroll tab honest.
function refreshApproved_(pairs) {
  const approved = approvedSet_(), done = {};
  pairs.forEach(p => {
    const key = p[0] + '|' + p[1];
    if (approved[key] && !done[key]) { done[key] = 1; writePayrollRow_(p[0], p[1]); }
  });
}

/* ---------------- Reminders ---------------- */

function dailyReminder() {
  const s = getSettings_();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const now = new Date();
  const dow = Utilities.formatDate(now, tz, 'EEE');
  const days = String(s.WorkDays || '').split(/[\s,]+/).map(d => d.slice(0, 3).toLowerCase());
  if (days.indexOf(dow.toLowerCase()) === -1) return;
  const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const logged = {};
  readEntries_().forEach(e => { if (e.date === today) logged[e.employee] = true; });
  getEmployees_().forEach(emp => {
    if (!emp.active || !emp.email || logged[emp.name]) return;
    MailApp.sendEmail({
      to: emp.email,
      subject: 'Log your hours for today',
      htmlBody: '<p>Hi ' + esc_(emp.name) + ',</p><p>No hours are logged for today yet. If you worked, take 30 seconds and enter them now.</p>' +
        (s.AppUrl ? '<p><a href="' + esc_(s.AppUrl) + '">Open the time app</a></p>' : ''),
    });
  });
}

function weeklySummary() {
  const s = getSettings_();
  if (!s.OwnerEmail) return;
  const ws = addDays_(mondayOf_(today_()), -7);
  const data = getWeek_({ role: 'owner' }, ws);
  let html = '<h2>Hours for the week of ' + ws + ' to ' + addDays_(ws, 6) + '</h2>' +
    '<table cellpadding="6" style="border-collapse:collapse" border="1"><tr><th>Employee</th><th>Days</th><th>Total hrs</th><th>OT hrs</th><th>Rate</th><th>Gross</th><th>Status</th></tr>';
  data.employees.forEach(emp => {
    const mins = emp.entries.reduce((t, e) => t + workMinutes_(e.start, e.end, e.breakMin), 0);
    const p = payFor_(mins, emp.rate, data.overtimeAfter, data.overtimeMultiplier);
    const days = {}; emp.entries.forEach(e => { days[e.date] = 1; });
    html += '<tr><td>' + esc_(emp.name) + '</td><td>' + Object.keys(days).length + '</td><td>' + p.totalHours.toFixed(2) +
      '</td><td>' + p.otHours.toFixed(2) + '</td><td>$' + emp.rate.toFixed(2) + '</td><td><b>$' + p.gross.toFixed(2) +
      '</b></td><td>' + (emp.approved ? 'Approved' : 'Not approved yet') + '</td></tr>';
  });
  html += '</table>' + (s.AppUrl ? '<p><a href="' + esc_(s.AppUrl) + '">Open the time app to review and approve</a></p>' : '');
  MailApp.sendEmail({ to: s.OwnerEmail, subject: 'Payroll hours: week of ' + ws, htmlBody: html });
}

function testEmail() {
  const s = getSettings_();
  if (!s.OwnerEmail) throw new Error('Put your email in Settings > OwnerEmail first');
  MailApp.sendEmail(s.OwnerEmail, 'Tightline Time test', 'Email from your time tracker works. Reminders go out at hour ' + s.ReminderHour + ' on ' + s.WorkDays + '.');
}

/* ---------------- Helpers ---------------- */

function cleanEntry_(e) {
  const id = String(e.id || '');
  if (!/^[A-Za-z0-9_-]{6,40}$/.test(id)) throw new Error('Bad entry id');
  const date = String(e.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Pick a date');
  if (date > addDays_(today_(), 1)) throw new Error('You cannot log time for a future date');
  const start = String(e.start || ''), end = String(e.end || '');
  if (!/^\d{2}:\d{2}$/.test(start) || !/^\d{2}:\d{2}$/.test(end)) throw new Error('Enter a start and end time');
  const breakMin = clampInt_(e.breakMin, 0, 600, 0);
  const mins = workMinutes_(start, end, breakMin);
  if (mins <= 0) throw new Error('Hours come out to zero. Check start, end, and break.');
  if (mins > MAX_SHIFT_MIN) throw new Error('That is more than 16 hours. Check the times.');
  return {
    id: id, date: date, start: start, end: end, breakMin: breakMin,
    hours: Math.round(mins / 60 * 100) / 100,
    job: String(e.job || '').trim().slice(0, 80),
    notes: String(e.notes || '').trim().slice(0, 500),
  };
}

function workMinutes_(start, end, breakMin) {
  const a = toMin_(start), b = toMin_(end);
  let span = b - a;
  if (span < 0) span += 1440; // shift ran past midnight
  return span - (Number(breakMin) || 0);
}

function payFor_(mins, rate, otAfter, otMult) {
  const reg = Math.min(mins, otAfter * 60);
  const ot = Math.max(0, mins - otAfter * 60);
  const gross = (reg / 60) * rate + (ot / 60) * rate * otMult;
  return {
    regHours: round2_(reg / 60), otHours: round2_(ot / 60), totalHours: round2_(mins / 60), gross: round2_(gross),
  };
}

function recentJobs_(entries) {
  const seen = {}, out = [];
  entries.slice().sort(byDateTime_).reverse().forEach(e => {
    const j = (e.job || '').trim();
    if (j && !seen[j.toLowerCase()] && out.length < 25) { seen[j.toLowerCase()] = 1; out.push(j); }
  });
  return out;
}

function byDateTime_(a, b) { return (a.date + a.start) < (b.date + b.start) ? -1 : (a.date + a.start) > (b.date + b.start) ? 1 : 0; }
function toMin_(t) { const p = String(t).split(':'); return Number(p[0]) * 60 + Number(p[1]); }
function round2_(n) { return Math.round(n * 100) / 100; }
function clampInt_(v, lo, hi, dflt) { const n = Math.round(Number(v)); return isFinite(n) ? Math.min(hi, Math.max(lo, n)) : dflt; }
function pad_(n) { return (n < 10 ? '0' : '') + n; }
function today_() { return Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd'); }
function stamp_() { return Utilities.formatDate(new Date(), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm'); }
function parseDate_(s) { const p = String(s).split('-').map(Number); return new Date(Date.UTC(p[0], p[1] - 1, p[2])); }
function fmtDate_(d) { return d.getUTCFullYear() + '-' + pad_(d.getUTCMonth() + 1) + '-' + pad_(d.getUTCDate()); }
function addDays_(s, n) { const d = parseDate_(s); d.setUTCDate(d.getUTCDate() + n); return fmtDate_(d); }
function mondayOf_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(s))) throw new Error('Bad week');
  const d = parseDate_(s); const back = (d.getUTCDay() + 6) % 7; d.setUTCDate(d.getUTCDate() - back); return fmtDate_(d);
}
function randomPin_() { return String(Math.floor(100000 + Math.random() * 900000)); }
// Stop the sheet from treating text like "=SUM()" or "-3" as a formula.
function escapeCell_(s) { return /^[=+\-@]/.test(s) ? "'" + s : s; }
function unescapeCell_(s) { return String(s || ''); }
function esc_(s) { return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
