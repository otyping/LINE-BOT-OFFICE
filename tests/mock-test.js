/**
 * ทดสอบ Code.gs ด้วย Node โดยจำลอง SpreadsheetApp / UrlFetchApp / LINE
 *   node tests/mock-test.js
 * ลำดับที่ทดสอบ: ย้ายข้อมูลจากชีทเดิม -> init -> แจ้งวันทำงาน (ลงทันที ไม่ต้องขออนุมัติ)
 *                -> เคลียร์คำขอเวรที่ค้างจากระบบเดิม -> ตรวจสิทธิ์ -> สร้างตารางรอบ
 *                -> โอที (แจ้งล่วงหน้า -> CEO อนุมัติ/ไม่อนุมัติ -> แจ้งย้อนหลัง -> ตารางโอที)
 */
const fs = require('fs');
const path = require('path');

/* ---------- จำลอง Spreadsheet ---------- */
function Sheet(name, grid) { this.name = name; this.g = grid || []; this.notes = {}; this.rules = []; this.merges = []; this.fmt = {}; }
Sheet.prototype.getName = function () { return this.name; };
Sheet.prototype.getLastRow = function () { return this.g.length; };
Sheet.prototype.getLastColumn = function () { return Math.max(0, ...this.g.map(r => (r || []).length)); };
Sheet.prototype.getRange = function (r, c, nr, nc) {
  if (typeof r === 'string') { // getRange('A:A') ใช้แค่จัดรูปแบบ จึงคืน no-op
    return { setNumberFormat() { return this; }, setDataValidation() { return this; }, setValues() { return this; } };
  }
  nr = nr || 1; nc = nc || 1;
  const s = this;
  const api = {
    getValues: () => Array.from({ length: nr }, (_, i) =>
      Array.from({ length: nc }, (_, j) => { const v = (s.g[r - 1 + i] || [])[c - 1 + j]; return v === undefined || v === null ? '' : v; })),
    getDisplayValues() { return this.getValues().map(row => row.map(v => v instanceof Date ? fmt(v) : String(v))); },
    getDisplayValue() { return this.getDisplayValues()[0][0]; },
    setValue(v) { (s.g[r - 1] = s.g[r - 1] || [])[c - 1] = v; return this; },
    setValues(vs) { vs.forEach((row, i) => row.forEach((v, j) => { (s.g[r - 1 + i] = s.g[r - 1 + i] || [])[c - 1 + j] = v === undefined ? '' : v; })); return this; },
    getNote: () => s.notes[r + ',' + c] || '',
    setNote(n) { s.notes[r + ',' + c] = n; return this; },
    clearNote() { s.notes = {}; return this; },
    merge() { s.merges.push({ c: c, nc: nc }); return this; },
    breakApart() { s.merges = s.merges.filter(m => m.c < c || m.c + m.nc - 1 > c + nc - 1); return this; }
  };
  api.setNumberFormat = f => { s.fmt[[r, c, nr, nc].join()] = f; return api; };
  ['setFontWeight', 'setDataValidation', 'setHorizontalAlignment',
   'setBackground', 'setBackgrounds', 'setFontColor', 'setFontSize'].forEach(m => (api[m] = () => api));
  return api;
};
Sheet.prototype.getDataRange = function () { return this.getRange(1, 1, Math.max(this.g.length, 1), Math.max(this.getLastColumn(), 1)); };
Sheet.prototype.appendRow = function (r) { this.g.push(r); };
Sheet.prototype.clear = function () { this.g = []; this.notes = {}; this.fmt = {}; return this; };
Sheet.prototype.getMaxRows = function () { return Math.max(this.g.length, 1000); };
Sheet.prototype.getMaxColumns = function () { return Math.max(this.getLastColumn(), 26); };
Sheet.prototype.clearConditionalFormatRules = function () { this.rules = []; };
Sheet.prototype.getConditionalFormatRules = function () { return this.rules.slice(); };
Sheet.prototype.setConditionalFormatRules = function (r) { this.rules = r; };
['setFrozenRows', 'setColumnWidth'].forEach(m => (Sheet.prototype[m] = function () { return this; }));
// Google Sheets จะ error ถ้าตรึงคอลัมน์ผ่ากลางเซลล์ที่ merge ไว้ จำลองไว้กันบั๊กซ้ำ
Sheet.prototype.setFrozenColumns = function (n) {
  if (n > 0 && this.merges.some(m => m.c <= n && m.c + m.nc - 1 > n)) {
    throw new Error("Sorry, you can't freeze columns which contain only part of a merged cell.");
  }
  return this;
};

const fmt = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const sheets = {};
const addSheet = (n, g) => (sheets[n] = new Sheet(n, g));

global.SpreadsheetApp = {
  getActiveSpreadsheet: () => ({
    getSheetByName: n => sheets[n] || null,
    insertSheet: n => addSheet(n, []),
    getSpreadsheetTimeZone: () => 'Asia/Bangkok'
  }),
  flush() {},
  newDataValidation: () => ({ requireValueInList() { return this; }, build() { return {}; } }),
  BooleanCriteria: { CUSTOM_FORMULA: 'CUSTOM_FORMULA' },
  newConditionalFormatRule: () => ({
    whenTextStartsWith() { return this; }, whenTextEqualTo() { return this; },
    whenFormulaSatisfied(f) { this.f = f; return this; },
    setBackground() { return this; }, setFontColor() { return this; }, setRanges() { return this; },
    build() {
      const f = this.f;
      return { getBooleanCondition: () => f ? { getCriteriaType: () => 'CUSTOM_FORMULA', getCriteriaValues: () => [f] } : null };
    }
  })
};
global.Utilities = {
  formatDate: (d, tz, f) => f === 'yyyy-MM-dd' ? fmt(d)
    : `${d.getDate()}/${d.getMonth() + 1} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`
};
global.LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
global.PropertiesService = { getScriptProperties: () => ({ getProperty: k => 'P_' + k }) };

const pushed = [];
let who = 'U1';
global.UrlFetchApp = {
  fetch: (u, o) => {
    if (u.includes('verify')) return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ sub: who, name: 'n' }) };
    pushed.push(JSON.parse(o.payload));
    return { getResponseCode: () => 200, getContentText: () => '' };
  }
};
global.ContentService = { createTextOutput: t => ({ t, setMimeType() { return this; } }), MimeType: { JSON: 1 } };

/* ---------- ชีทตั้งต้น: ชีทเดิมแนวกว้าง + ผู้ใช้ระบบ ---------- */
const today = new Date();
const dayOf = n => new Date(today.getFullYear(), today.getMonth(), today.getDate() + n);
const D = n => fmt(dayOf(n));

// แถว 2 มีวันที่ 2 บล็อกคั่นด้วยช่องว่าง เพื่อทดสอบว่า migrate ข้ามช่องว่างได้
const legacy = [
  ['รายชื่อพนักงาน'],
  ['รหัส', 'ชื่อ - สกุล', 'ชื่อในกลุ่มไลน์', 'แผนก', 'ชื่อเล่น', dayOf(-2), dayOf(-1), '', dayOf(30)],
  [],
  ['KB028', 'น.ส.ฐิติมา  นิลสีอ่อน', '', 'ทีมปลูก', 'ทราย', '8.00-17.00', 'หยุด', '', '8.00-17.00'],
  ['013', 'น.ส.บังออล ปานยิ้ม', '', 'ทีมปลูก', 'น้าน้อง', '8.00-17.00'],
  ['', 'น.ส.น้ำค้าง ศรี', '', 'ห้องทริม', ''],
  ['หมายเหตุ']
];
addSheet('รหัสพนักงาน', legacy);
addSheet('ผู้ใช้ระบบ', [['h'], ['U1', 'หัวหน้า', 'หัวหน้างาน', ''], ['UHR', 'พี่HR', 'HR', ''], ['UCEO', 'บอส', 'CEO', '']]);

/* ---------- โหลด Code.gs ---------- */
eval(fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8') +
  ';global.handleApi_=handleApi_;global.checkSetup=checkSetup;global.handleLine_=handleLine_;' +
  'global.migrateFromLegacy=migrateFromLegacy;global.buildCurrentPeriod=buildCurrentPeriod;' +
  'global.period_=period_;global.scheduleText_=scheduleText_;global.CFG=CFG;global.buildCurrentOt=buildCurrentOt;');

/* ---------- ทดสอบ ---------- */
let fails = 0;
function ok(cond, label, extra) {
  console.log((cond ? 'ผ่าน  ' : 'ตก    ') + label + (extra !== undefined ? '  -> ' + JSON.stringify(extra) : ''));
  if (!cond) fails++;
}
const logRows = () => sheets['บันทึกเวร'].g.slice(1).filter(r => r && r[1]);
const cellOf = (ymd, code, status) => {
  const r = logRows().filter(x => fmt(new Date(x[0])) === ymd || x[0] === ymd).filter(x => x[1] === code)
    .filter(x => !status || x[4] === status).pop();
  return r ? r[3] : null;
};

console.log('=== ย้ายข้อมูลจากชีทเดิม ===');
console.log(migrateFromLegacy());
ok(sheets['พนักงาน'].g.length === 3, 'พนักงานถูกย้าย 2 คน (คนไม่มีรหัสถูกข้าม)', sheets['พนักงาน'].g.length - 1);
ok(logRows().length === 4, 'บันทึกเวรได้ 4 แถว รวมวันที่หลังช่องว่าง', logRows().length);
ok(cellOf(D(30), 'KB028', 'อนุมัติ') === '8.00-17.00', 'อ่านวันที่ในบล็อกหลังช่องว่างได้');
const before = logRows().length;
migrateFromLegacy();
ok(logRows().length === before, 'รันย้ำแล้วไม่เพิ่มข้อมูลซ้ำ', logRows().length);

console.log('\n=== ตรวจการตั้งค่า ===');
console.log(checkSetup());
const empRules = sheets['พนักงาน'].rules;
ok(empRules.length === 1 && empRules[0].getBooleanCondition().getCriteriaValues()[0].includes('TRIM($A2)'),
  'ชีทพนักงานมีกฎสีแดงสำหรับรหัสซ้ำ กฎเดียวแม้รัน setup หลายรอบ', empRules.length);
sheets['พนักงาน'].g.push([' 013 ', 'คนรหัสซ้ำ']);
ok(checkSetup().includes('รหัสพนักงานซ้ำ: 013 มี 2 แถว'), 'ตรวจการตั้งค่าเตือนเมื่อรหัสพนักงานซ้ำ');
sheets['พนักงาน'].g.pop();

console.log('\n=== หัวหน้างานแจ้งวันทำงาน (ไม่ต้องขออนุมัติ) ===');
let r = handleApi_({ action: 'init', idToken: 'x' });
ok(r.ok && r.employees.length === 2, 'init คืนพนักงาน 2 คน', r.employees && r.employees.length);
ok(r.dates.some(d => d.ymd === D(1)), 'ช่วงวันที่มีพรุ่งนี้');
ok(r.dates.some(d => d.ymd === D(-14)) && !r.dates.some(d => d.ymd === D(-15)), 'แจ้งย้อนหลังได้ 14 วัน');
ok(r.defaultDate === D(1), 'ค่าเริ่มต้นคือพรุ่งนี้', r.defaultDate);

let nPush = pushed.length;
r = handleApi_({ action: 'submit', idToken: 'x', date: D(1),
  groups: [{ shift: '21:00-05:00', codes: ['KB028'] }, { shift: '19.00-03.00', codes: ['013'] }] });
ok(r.ok && !r.id && r.saved && r.saved.id === 'R0001', 'แจ้งล่วงหน้า ลงตารางทันที ไม่มีคำขอรออนุมัติ', r);
ok(cellOf(D(1), 'KB028', 'อนุมัติ') === '21.00-05.00' && cellOf(D(1), 'KB028', 'รออนุมัติ') === null, 'เวลาถูกจัดรูปแบบเป็น 21.00-05.00 และเป็นอนุมัติเลย');
ok(logRows().some(x => x[5] === 'R0001' && x[6] === 'หัวหน้า' && x[7] === CFG.AUTO_BY), 'บันทึกชื่อผู้แจ้ง ผู้พิจารณาระบุว่าอัตโนมัติ');
ok(scheduleText_(1).includes('21.00-05.00') && !scheduleText_(1).includes('(รอ)'), 'ดูตารางพรุ่งนี้เห็นค่าทันที ไม่มี (รอ)');

console.log('\n=== แก้เวรเดิมของวันเดียวกัน ===');
r = handleApi_({ action: 'submit', idToken: 'x', date: D(1), groups: [{ shift: '8.00-17.00', codes: ['KB028'] }] });
ok(r.saved && !r.id && cellOf(D(1), 'KB028', 'อนุมัติ') === '8.00-17.00', 'ค่าใหม่ทับค่าเดิมทันที', r);
ok(logRows().filter(x => x[0] === D(1) && x[1] === 'KB028' && x[4] === 'อนุมัติ').length === 1 &&
   logRows().filter(x => x[0] === D(1) && x[1] === 'KB028' && x[4] === 'แทนที่').length === 1, 'แถวเดิมเป็นแทนที่ เหลืออนุมัติแถวเดียว');
ok(cellOf(D(1), '013', 'อนุมัติ') === '19.00-03.00', 'คนอื่นในวันเดียวกันไม่ถูกกระทบ');

console.log('\n=== แจ้งย้อนหลัง ===');
r = handleApi_({ action: 'submit', idToken: 'x', date: D(-5), groups: [{ shift: '9.00-18.00', codes: ['KB028', '013'] }] });
ok(r.ok && r.saved && !r.id && r.saved.groups[0].names.length === 2, 'วันย้อนหลัง ลงตารางทันที', r);
r = handleApi_({ action: 'submit', idToken: 'x', date: D(-5), groups: [{ shift: 'หยุด', codes: ['KB028'] }] });
ok(r.saved && !r.id && cellOf(D(-5), 'KB028', 'อนุมัติ') === 'หยุด', 'แก้เวรย้อนหลังที่ลงไว้แล้ว ก็ไม่ต้องรอใคร', r);
r = handleApi_({ action: 'submit', idToken: 'x', date: D(0), groups: [{ shift: '8.00-17.00', codes: ['013'] }] });
ok(r.saved && !r.id, 'วันนี้ก็ลงทันที', r);
ok(!!handleApi_({ action: 'submit', idToken: 'x', date: D(-15), groups: [{ shift: '8.00-17.00', codes: ['013'] }] }).error, 'ย้อนหลังเกิน 14 วันแจ้งไม่ได้');
ok(pushed.length === nPush, 'แจ้งวันทำงานไม่ push หาใครเลย');

console.log('\n=== HR แจ้งเอง ===');
who = 'UHR';
r = handleApi_({ action: 'submit', idToken: 'x', date: D(3), groups: [{ shift: '9.00-18.00', codes: ['KB028'] }] });
ok(r.ok && r.saved && !r.id, 'HR แจ้งล่วงหน้า ลงตารางทันที', r);
ok(cellOf(D(3), 'KB028', 'อนุมัติ') === '9.00-18.00' && logRows().some(x => x[5] === r.saved.id && x[7] === 'พี่HR'), 'ผู้พิจารณาคือ HR คนที่แจ้ง');

console.log('\n=== คำขอเวรที่ค้างจากระบบเดิม ===');
// จำลองคำขอที่ส่งไว้ก่อนยกเลิกขั้นตอนอนุมัติ ซึ่งยังต้องเคลียร์ได้
const legacyReq = (id, ymd, code, label, shift) => {
  sheets['บันทึกเวร'].g.push([ymd, code, label, shift, 'รออนุมัติ', id, 'หัวหน้า', '', new Date()]);
  sheets['คำขอ'].g.push([id, new Date(), 'U1', 'หัวหน้า', ymd, shift + ': ' + label,
    JSON.stringify([{ code: code, label: label, shift: shift }]), '[]', 'รออนุมัติ', '', '', '']);
};
legacyReq('R0901', D(4), '013', 'น้าน้อง', '10.00-18.00');
legacyReq('R0902', D(5), 'KB028', 'ทราย', '10.00-18.00');
legacyReq('R0903', D(6), '013', 'น้าน้อง', '10.00-18.00');
r = handleApi_({ action: 'init', idToken: 'x' });
ok(r.pending && r.pending.length === 3, 'HR ยังเห็นคำขอที่ค้างอยู่', r.pending && r.pending.length);
ok(scheduleText_(4).includes('(รอ) 10.00-18.00'), 'ตารางยังแสดง (รอ) ของคำขอที่ค้าง');
r = handleApi_({ action: 'approve', idToken: 'x', id: 'R0901', edits: { '013': 'หยุด' } });
ok(r.ok && !r.skipped.length && cellOf(D(4), '013', 'อนุมัติ') === 'หยุด', 'HR อนุมัติพร้อมแก้เวลาได้', r);
r = handleApi_({ action: 'reject', idToken: 'x', id: 'R0902', reason: 'ยังไม่ยืนยัน' });
ok(r.ok && cellOf(D(5), 'KB028', 'อนุมัติ') === null && cellOf(D(5), 'KB028', 'ไม่อนุมัติ') === '10.00-18.00', 'ไม่อนุมัติได้ แถวเปลี่ยนเป็นไม่อนุมัติ ไม่ต้องคืนค่า', r);
who = 'U1';
r = handleApi_({ action: 'submit', idToken: 'x', date: D(6), groups: [{ shift: '8.00-17.00', codes: ['013'] }] });
who = 'UHR';
ok(cellOf(D(6), '013', 'อนุมัติ') === '8.00-17.00' && cellOf(D(6), '013', 'รออนุมัติ') === null &&
   handleApi_({ action: 'pending', idToken: 'x' }).pending.length === 0, 'แจ้งใหม่ทับคำขอที่ค้าง คำขอค้างถูกปิดเป็นแทนที่');

console.log('\n=== ตรวจสิทธิ์ ===');
who = 'U1';
ok(handleApi_({ action: 'approve', idToken: 'x', id: 'R0902' }).error === 'เมนูนี้สำหรับ HR เท่านั้น', 'หัวหน้างานอนุมัติไม่ได้');
who = 'UCEO';
ok(!!handleApi_({ action: 'submit', idToken: 'x', date: D(1), groups: [{ shift: '8.00-17.00', codes: ['013'] }] }).error, 'CEO แจ้งวันทำงานไม่ได้ (ใช้ได้เฉพาะโอที)');
who = 'UX';
ok(handleApi_({ action: 'submit', idToken: 'x', date: D(1), groups: [] }).error === 'บัญชีนี้ยังไม่มีสิทธิ์ใช้งาน', 'คนนอกใช้งานไม่ได้');

console.log('\n=== พนักงานลาออก ===');
who = 'U1';
sheets['พนักงาน'].g[1][6] = D(-1); // KB028 ออกเมื่อวาน
r = handleApi_({ action: 'init', idToken: 'x' });
ok(r.employees.length === 1 && r.employees[0].code === '013', 'คนลาออกไม่ขึ้นในฟอร์ม', r.employees.map(e => e.code));
ok(logRows().some(x => x[1] === 'KB028'), 'ประวัติของคนลาออกยังอยู่ครบ');
sheets['พนักงาน'].g[1][6] = '';

console.log('\n=== สร้างตารางรอบ ===');
buildCurrentPeriod();
const p = period_(fmt(today));
const view = sheets[CFG.VIEW_PREFIX + p.key];
ok(!!view, 'สร้างชีท ' + CFG.VIEW_PREFIX + p.key);
ok(view.g[1].length >= 4 + 28 && view.g[1].length <= 4 + 31, 'หัวตารางมีคอลัมน์วันครบทั้งรอบ', view.g[1].length - 4);
ok(view.g[1].indexOf('ชื่อในกลุ่มไลน์') < 0 && view.g[1].slice(0, 4).join() === 'รหัส,ชื่อ - สกุล,แผนก,ชื่อเล่น', 'ชีทรอบไม่มีคอลัมน์ชื่อในกลุ่มไลน์', view.g[1].slice(0, 4));
ok(view.fmt[[4, 1, 2, 1].join()] === '@' && view.g.slice(3).some(x => x && x[0] === '013'), 'คอลัมน์รหัสเป็นข้อความ รหัส 013 ไม่กลายเป็น 13');
ok(view.g[1].indexOf('ค่าแรง') < 0 && sheets['พนักงาน'].g[0].indexOf('ค่าแรง') < 0, 'ไม่มีคอลัมน์ค่าแรงในชีทรอบและชีทพนักงาน');
ok(view.g.slice(3).some(row => row && row.indexOf('8.00-17.00') >= 0), 'ตารางดึงค่าจากบันทึกเวรมาแสดง');
buildCurrentPeriod();
ok(sheets[CFG.VIEW_PREFIX + p.key].g.slice(3).filter(x => x && x[0]).length === 2, 'สร้างซ้ำแล้วไม่มีแถวค้าง');

console.log('\n=== โอที: หัวหน้างานแจ้งล่วงหน้า รอ CEO ===');
const otRows = () => sheets['บันทึกโอที'].g.slice(1).filter(x => x && x[1]);
const otOf = (ymd, code, status) => otRows().filter(x => x[0] === ymd && x[1] === code && (!status || x[5] === status)).pop() || null;
who = 'U1';
nPush = pushed.length;
r = handleApi_({ action: 'otSubmit', idToken: 'x', date: D(1),
  groups: [{ hours: 3, codes: ['KB028'] }, { hours: 2.5, holiday: true, codes: ['013'] }] });
ok(r.ok && r.id === 'T0001' && !r.saved, 'ส่งคำขอโอทีสำเร็จ เลขคำขอขึ้นต้นด้วย T', r);
ok(otOf(D(1), 'KB028', 'รออนุมัติ')[3] === 3 && otOf(D(1), '013', 'รออนุมัติ')[4] === 'ใช่', 'บันทึกโอทีเป็นรออนุมัติ พร้อมเครื่องหมายทำงานวันหยุด');
ok(pushed.length === nPush + 1 && pushed[pushed.length - 1].to === 'UCEO', 'push หา CEO ไม่ใช่ HR', pushed[pushed.length - 1].to);
ok(handleApi_({ action: 'otApprove', idToken: 'x', id: 'T0001' }).error === 'เมนูนี้สำหรับ CEO เท่านั้น', 'หัวหน้างานอนุมัติโอทีไม่ได้');
who = 'UHR';
ok(handleApi_({ action: 'otApprove', idToken: 'x', id: 'T0001' }).error === 'เมนูนี้สำหรับ CEO เท่านั้น', 'HR อนุมัติโอทีไม่ได้');
ok(!!handleApi_({ action: 'otSubmit', idToken: 'x', date: D(1), groups: [{ hours: 2.3, codes: ['KB028'] }] }).error, 'ชั่วโมงที่ไม่ลงครึ่งชั่วโมง แจ้งไม่ได้');
ok(!!handleApi_({ action: 'otSubmit', idToken: 'x', date: D(1), groups: [{ hours: 0, codes: ['KB028'] }] }).error, '0 ชั่วโมงและไม่ใช่วันหยุด แจ้งไม่ได้');

console.log('\n=== โอที: CEO อนุมัติพร้อมแก้ชั่วโมง ===');
who = 'UCEO';
r = handleApi_({ action: 'init', idToken: 'x' });
ok(r.otPending && r.otPending.length === 1 && r.otPending[0].entries.length === 2, 'CEO เห็นคำขอโอทีค้าง 1 ใบ', r.otPending && r.otPending.length);
nPush = pushed.length;
r = handleApi_({ action: 'otApprove', idToken: 'x', id: 'T0001', edits: { KB028: '4' } });
ok(r.ok && !r.skipped.length, 'อนุมัติโอทีสำเร็จ', r);
ok(otOf(D(1), 'KB028', 'อนุมัติ')[3] === 4 && otOf(D(1), 'KB028', 'อนุมัติ')[8] === 'บอส', 'CEO แก้ชั่วโมงเป็น 4 และเป็นผู้พิจารณา');
ok(pushed.length === nPush + 1 && pushed[pushed.length - 1].to === 'U1', 'แจ้งผลกลับผู้แจ้ง');
ok(!!handleApi_({ action: 'otApprove', idToken: 'x', id: 'T0001' }).error, 'อนุมัติซ้ำไม่ได้');

console.log('\n=== โอที: แจ้งย้อนหลัง ถือว่าอนุมัติแล้ว ===');
who = 'U1';
nPush = pushed.length;
r = handleApi_({ action: 'otSubmit', idToken: 'x', date: D(-3), groups: [{ hours: 5, codes: ['KB028', '013'] }] });
ok(r.ok && r.saved && !r.id, 'วันย้อนหลัง บันทึกทันที ไม่มีคำขอรออนุมัติ', r);
ok(otOf(D(-3), 'KB028', 'อนุมัติ')[3] === 5 && otOf(D(-3), 'KB028', 'อนุมัติ')[8] === CFG.AUTO_BY, 'ผู้พิจารณาระบุว่าอัตโนมัติ');
ok(pushed.length === nPush, 'แจ้งย้อนหลังครั้งแรกไม่ push หาใคร');
r = handleApi_({ action: 'otSubmit', idToken: 'x', date: D(-3), groups: [{ hours: 6, codes: ['KB028'] }] });
ok(r.id && !r.saved && otOf(D(-3), 'KB028', 'อนุมัติ')[3] === 5, 'แก้โอทีที่อนุมัติแล้ว ต้องรอ CEO ค่าเดิมยังอยู่', r);
ok(r.warnings.length === 1 && r.warnings[0].includes('อนุมัติไว้แล้ว 5 ชม.') && r.warnings[0].includes('หัวหน้า ขอแก้เป็น 6 ชม.'), 'คำเตือนบอกค่าเดิม ค่าใหม่ และชื่อคนแก้', r.warnings);
const card = pushed[pushed.length - 1];
ok(pushed.length === nPush + 1 && card.to === 'UCEO' && JSON.stringify(card.messages).includes('ขอแก้โอทีที่อนุมัติแล้ว') &&
   JSON.stringify(card.messages).includes('หัวหน้า ขอแก้เป็น 6 ชม.'), 'แจ้งเตือน CEO พร้อมรายละเอียดการแก้');
who = 'UCEO';
ok(handleApi_({ action: 'otPending', idToken: 'x' }).pending.some(x => x.id === r.id && x.warnings.length === 1 && x.entries[0].old === '5 ชม.'), 'หน้าอนุมัติของ CEO เห็นค่าเดิมและคำเตือน');
handleApi_({ action: 'otApprove', idToken: 'x', id: r.id });
ok(otOf(D(-3), 'KB028', 'อนุมัติ')[3] === 6 &&
   otRows().filter(x => x[0] === D(-3) && x[1] === 'KB028' && x[5] === 'แทนที่').length === 1, 'CEO อนุมัติแล้วค่าใหม่จึงแทนที่ค่าเดิม');
who = 'U1';
handleApi_({ action: 'otSubmit', idToken: 'x', date: D(-4), groups: [{ hours: 2, codes: ['013'] }] });
r = handleApi_({ action: 'otSubmit', idToken: 'x', date: D(-4), groups: [{ hours: 3, codes: ['KB028', '013'] }] });
ok(r.saved && r.id && r.saved.groups[0].names.join() === 'ทราย' && r.groups[0].names.join() === 'น้าน้อง', 'คำขอผสม: คนใหม่ลงทันที คนที่แก้ของเดิมรอ CEO', r);
r = handleApi_({ action: 'otSubmit', idToken: 'x', date: D(0), groups: [{ hours: 2, codes: ['013'] }] });
ok(r.id && !r.saved, 'วันนี้ไม่นับเป็นย้อนหลัง ต้องรอ CEO', r);
const todayReq = r.id;

console.log('\n=== โอที: ไม่อนุมัติ และ CEO แจ้งเอง ===');
who = 'UCEO';
r = handleApi_({ action: 'otReject', idToken: 'x', id: todayReq, reason: 'ไม่จำเป็น' });
ok(r.ok && otOf(D(0), '013', 'ไม่อนุมัติ') && !otOf(D(0), '013', 'อนุมัติ'), 'ไม่อนุมัติสำเร็จ ไม่มีโอทีในวันนั้น', r);
who = 'U1';
r = handleApi_({ action: 'otSubmit', idToken: 'x', date: D(2), groups: [{ hours: 2, codes: ['KB028'] }] });
const waitReq = r.id;
who = 'UCEO';
nPush = pushed.length;
r = handleApi_({ action: 'otSubmit', idToken: 'x', date: D(2), groups: [{ hours: 3, codes: ['KB028'] }] });
ok(r.saved && !r.id && otOf(D(2), 'KB028', 'อนุมัติ')[3] === 3 && pushed.length === nPush, 'CEO แจ้งเอง ลงทันที ไม่ push', r);
ok(!otOf(D(2), 'KB028', 'รออนุมัติ') && handleApi_({ action: 'otPending', idToken: 'x' }).pending.every(x => x.id !== waitReq),
  'คำขอที่ค้างของคน-วันเดียวกันถูกปิดเป็นแทนที่');

console.log('\n=== สร้างตารางโอทีรอบ ===');
buildCurrentOt();
const op = period_(fmt(today));
const otView = sheets[CFG.OT_VIEW_PREFIX + op.key];
ok(!!otView, 'สร้างชีท ' + CFG.OT_VIEW_PREFIX + op.key);
const nDays = otView.g[1].length - 6;
ok(otView.g[1].indexOf('ชื่อในกลุ่มไลน์') < 0 && otView.fmt[[4, 1, 2, 1].join()] === '@', 'ชีทโอทีไม่มีคอลัมน์ชื่อในกลุ่มไลน์ และรหัสเป็นข้อความ');
ok(otView.g[1][otView.g[1].length - 2] === 'รวม OT' && nDays >= 28 && nDays <= 31, 'ทั้งรอบอยู่ในตารางเดียว มีคอลัมน์รวม OT ต่อท้าย', nDays);
const wantSum = otRows().filter(x => x[1] === 'KB028' && x[5] === 'อนุมัติ' && x[0] >= op.start && x[0] <= op.end).reduce((a, x) => a + x[3], 0);
const kbRow = otView.g.slice(3).find(x => x && x[0] === 'KB028');
ok(kbRow && kbRow[kbRow.length - 2] === wantSum && wantSum > 0, 'รวม OT นับเฉพาะรายการที่อนุมัติแล้ว', wantSum);
buildCurrentOt();
ok(sheets[CFG.OT_VIEW_PREFIX + op.key].g.slice(3).filter(x => x && /^(KB028|013)$/.test(x[0])).length === 2, 'สร้างซ้ำแล้วไม่มีแถวค้าง');

console.log('\n=== การแจ้งเตือน LINE ===');
ok(pushed.length >= 3, 'มีการ push หา CEO และผู้แจ้ง', pushed.length);
ok(pushed.every(m => m.to && m.to[0] === 'U'), 'push เฉพาะแชทส่วนตัว ไม่ push เข้ากลุ่ม');

console.log('\n=== เรียกบอทในกลุ่มด้วยชื่อ ===');
const say = (text, srcType) => {
  const n = pushed.length;
  handleLine_([{ type: 'message', replyToken: 'rt', source: { type: srcType || 'group', userId: 'U1' },
    message: { type: 'text', text: text } }]);
  return pushed.slice(n).map(m => (m.messages || [])[0]);
};
ok(say('จำปี')[0] && say('จำปี')[0].type === 'flex', 'พิมพ์ชื่อบอทเฉย ๆ ได้เมนูปุ่มกด');
ok(say('แจ้งงาน').length === 0, 'คำเดิม "แจ้งงาน" ไม่เรียกบอทแล้ว');
ok(say('จำปี เมนู')[0].type === 'flex', 'สั่ง "จำปี เมนู" ได้เมนู');
ok(say('จำปี ตารางวันนี้')[0].text.indexOf('ตาราง') === 0, 'สั่งงานต่อท้ายชื่อบอทได้');
ok(!!say('จำปี ตารางวันนี้')[0].quickReply, 'คำตอบมีปุ่มลัดให้กดต่อ');
ok(say('วันนี้กินอะไรดี').length === 0, 'ข้อความอื่นในกลุ่ม บอทเงียบ');
ok(say('อะไรก็ได้', 'user').length === 1, 'แชทส่วนตัวตอบเสมอ');
ok(say('บันทึกโอที ส. 10 ต.ค. 69 (#T0001)\n3 ชม.: ทราย', 'user').length === 0 &&
   say('ขออนุมัติโอที ส. 10 ต.ค. 69 (#T0002)\n3 ชม.: ทราย', 'user').length === 0 &&
   say('บันทึกตาราง ส. 10 ต.ค. 69 (#R0001)\n8.00-17.00: ทราย', 'user').length === 0, 'ข้อความสรุปที่ฟอร์มส่งเองในแชทส่วนตัว บอทไม่ตอบ');

console.log('\n' + (fails ? 'มีข้อทดสอบไม่ผ่าน ' + fails + ' ข้อ' : 'ผ่านทั้งหมด'));
process.exit(fails ? 1 : 0);
