/**
 * ทดสอบ Code.gs ด้วย Node โดยจำลอง SpreadsheetApp / UrlFetchApp / LINE
 *   node tests/mock-test.js
 * ลำดับที่ทดสอบ: ย้ายข้อมูลจากชีทเดิม -> init -> submit -> approve (แก้เวลา)
 *                -> submit -> reject -> ตรวจสิทธิ์ -> สร้างตารางรอบ
 */
const fs = require('fs');
const path = require('path');

/* ---------- จำลอง Spreadsheet ---------- */
function Sheet(name, grid) { this.name = name; this.g = grid || []; this.notes = {}; this.rules = []; }
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
    clearNote() { s.notes = {}; return this; }
  };
  ['setFontWeight', 'setNumberFormat', 'setDataValidation', 'merge', 'setHorizontalAlignment',
   'setBackground', 'setFontColor'].forEach(m => (api[m] = () => api));
  return api;
};
Sheet.prototype.getDataRange = function () { return this.getRange(1, 1, Math.max(this.g.length, 1), Math.max(this.getLastColumn(), 1)); };
Sheet.prototype.appendRow = function (r) { this.g.push(r); };
Sheet.prototype.clear = function () { this.g = []; this.notes = {}; return this; };
Sheet.prototype.clearConditionalFormatRules = function () { this.rules = []; };
Sheet.prototype.setConditionalFormatRules = function (r) { this.rules = r; };
['setFrozenRows', 'setFrozenColumns', 'setColumnWidth'].forEach(m => (Sheet.prototype[m] = function () { return this; }));

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
  newConditionalFormatRule: () => ({
    whenTextStartsWith() { return this; }, whenTextEqualTo() { return this; },
    setBackground() { return this; }, setFontColor() { return this; }, setRanges() { return this; },
    build() { return {}; }
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
  ['รหัส', 'ชื่อ - สกุล', 'ชื่อในกลุ่มไลน์', 'แผนก', 'ชื่อเล่น', '30,000.00', 'เดือน', dayOf(-2), dayOf(-1), '', dayOf(30)],
  [],
  ['KB028', 'น.ส.ฐิติมา  นิลสีอ่อน', '', 'ทีมปลูก', 'ทราย', '13,000', 'เดือน', '8.00-17.00', 'หยุด', '', '8.00-17.00'],
  ['013', 'น.ส.บังออล ปานยิ้ม', '', 'ทีมปลูก', 'น้าน้อง', '15,000', 'เดือน', '8.00-17.00'],
  ['', 'น.ส.น้ำค้าง ศรี', '', 'ห้องทริม', ''],
  ['หมายเหตุ']
];
addSheet('รหัสพนักงาน', legacy);
addSheet('ผู้ใช้ระบบ', [['h'], ['U1', 'หัวหน้า', 'หัวหน้างาน', ''], ['UHR', 'พี่HR', 'HR', '']]);

/* ---------- โหลด Code.gs ---------- */
eval(fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8') +
  ';global.handleApi_=handleApi_;global.checkSetup=checkSetup;global.handleLine_=handleLine_;' +
  'global.migrateFromLegacy=migrateFromLegacy;global.buildCurrentPeriod=buildCurrentPeriod;' +
  'global.period_=period_;global.scheduleText_=scheduleText_;global.CFG=CFG;');

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

console.log('\n=== หัวหน้างานส่งคำขอ ===');
let r = handleApi_({ action: 'init', idToken: 'x' });
ok(r.ok && r.employees.length === 2, 'init คืนพนักงาน 2 คน', r.employees && r.employees.length);
ok(r.dates.some(d => d.ymd === D(1)), 'ช่วงวันที่มีพรุ่งนี้');
ok(r.dates.some(d => d.ymd === D(-3)), 'แจ้งย้อนหลังได้ 7 วัน');
ok(r.defaultDate === D(1), 'ค่าเริ่มต้นคือพรุ่งนี้', r.defaultDate);

r = handleApi_({ action: 'submit', idToken: 'x', date: D(1),
  groups: [{ shift: '21:00-05:00', codes: ['KB028'] }, { shift: '19.00-03.00', codes: ['013'] }] });
ok(r.ok && r.id === 'R0001', 'ส่งคำขอสำเร็จ', r);
ok(cellOf(D(1), 'KB028', 'รออนุมัติ') === '21.00-05.00', 'เวลาถูกจัดรูปแบบเป็น 21.00-05.00');
ok(scheduleText_(1).includes('(รอ) 21.00-05.00'), 'ดูตารางพรุ่งนี้เห็นสถานะรออนุมัติ');

console.log('\n=== HR อนุมัติพร้อมแก้เวลา ===');
who = 'UHR';
r = handleApi_({ action: 'init', idToken: 'x' });
ok(r.pending && r.pending.length === 1, 'HR เห็นคำขอค้าง 1 รายการ', r.pending && r.pending.length);
r = handleApi_({ action: 'approve', idToken: 'x', id: 'R0001', edits: { '013': 'หยุด' } });
ok(r.ok && !r.skipped.length, 'อนุมัติสำเร็จ', r);
ok(cellOf(D(1), '013', 'อนุมัติ') === 'หยุด', 'HR แก้เวลาของ น้าน้อง เป็น หยุด');
ok(scheduleText_(1).includes('21.00-05.00') && !scheduleText_(1).includes('(รอ)'), 'ตารางพรุ่งนี้เป็นค่าอนุมัติแล้ว');

console.log('\n=== แทนที่เวรเดิมของวันเดียวกัน ===');
r = handleApi_({ action: 'submit', idToken: 'x', date: D(1), groups: [{ shift: '8.00-17.00', codes: ['KB028'] }] });
ok(r.warnings.length === 0, 'ไม่มีคำขออื่นค้างอยู่');
handleApi_({ action: 'approve', idToken: 'x', id: 'R0002' });
ok(cellOf(D(1), 'KB028', 'อนุมัติ') === '8.00-17.00', 'ค่าใหม่ทับค่าเดิม');
ok(logRows().filter(x => x[1] === 'KB028' && x[4] === 'แทนที่').length === 1, 'แถวเดิมถูกทำเครื่องหมายว่าแทนที่');

console.log('\n=== ไม่อนุมัติ ===');
r = handleApi_({ action: 'submit', idToken: 'x', date: D(2), groups: [{ shift: '10.00-18.00', codes: ['KB028'] }] });
r = handleApi_({ action: 'reject', idToken: 'x', id: 'R0003', reason: 'ยังไม่ยืนยัน' });
ok(r.ok, 'ไม่อนุมัติสำเร็จ', r);
ok(cellOf(D(2), 'KB028', 'อนุมัติ') === null, 'ไม่มีเวรที่อนุมัติในวันนั้น');
ok(cellOf(D(2), 'KB028', 'ไม่อนุมัติ') === '10.00-18.00', 'แถวเดิมเปลี่ยนเป็นไม่อนุมัติ ไม่ต้องคืนค่า');

console.log('\n=== ตรวจสิทธิ์ ===');
who = 'U1';
ok(handleApi_({ action: 'approve', idToken: 'x', id: 'R0003' }).error === 'เมนูนี้สำหรับ HR เท่านั้น', 'หัวหน้างานอนุมัติไม่ได้');
who = 'UX';
ok(handleApi_({ action: 'submit', idToken: 'x', date: D(1), groups: [] }).error === 'บัญชีนี้ยังไม่มีสิทธิ์ใช้งาน', 'คนนอกใช้งานไม่ได้');

console.log('\n=== พนักงานลาออก ===');
who = 'U1';
sheets['พนักงาน'].g[1][8] = D(-1); // KB028 ออกเมื่อวาน
r = handleApi_({ action: 'init', idToken: 'x' });
ok(r.employees.length === 1 && r.employees[0].code === '013', 'คนลาออกไม่ขึ้นในฟอร์ม', r.employees.map(e => e.code));
ok(logRows().some(x => x[1] === 'KB028'), 'ประวัติของคนลาออกยังอยู่ครบ');
sheets['พนักงาน'].g[1][8] = '';

console.log('\n=== สร้างตารางรอบ ===');
buildCurrentPeriod();
const p = period_(fmt(today));
const view = sheets[CFG.VIEW_PREFIX + p.key];
ok(!!view, 'สร้างชีท ' + CFG.VIEW_PREFIX + p.key);
ok(view.g[1].length === 7 + 31 || view.g[1].length === 7 + 30 || view.g[1].length === 7 + 29,
  'หัวตารางมีคอลัมน์วันครบทั้งรอบ', view.g[1].length - 7);
ok(view.g.slice(3).some(row => row && row.indexOf('8.00-17.00') >= 0), 'ตารางดึงค่าจากบันทึกเวรมาแสดง');
buildCurrentPeriod();
ok(sheets[CFG.VIEW_PREFIX + p.key].g.slice(3).filter(x => x && x[0]).length === 2, 'สร้างซ้ำแล้วไม่มีแถวค้าง');

console.log('\n=== การแจ้งเตือน LINE ===');
ok(pushed.length >= 3, 'มีการ push หา HR และผู้แจ้ง', pushed.length);
ok(pushed.every(m => m.to && m.to[0] === 'U'), 'push เฉพาะแชทส่วนตัว ไม่ push เข้ากลุ่ม');

console.log('\n=== เรียกบอทในกลุ่มด้วยชื่อ ===');
const say = (text, srcType) => {
  const n = pushed.length;
  handleLine_([{ type: 'message', replyToken: 'rt', source: { type: srcType || 'group', userId: 'U1' },
    message: { type: 'text', text: text } }]);
  return pushed.slice(n).map(m => (m.messages || [])[0]);
};
ok(say('จำปี')[0] && say('จำปี')[0].type === 'flex', 'พิมพ์ชื่อบอทเฉย ๆ ได้เมนูปุ่มกด');
ok(say('แจ้งงาน')[0] && say('แจ้งงาน')[0].type === 'flex', 'คำเดิม "แจ้งงาน" ยังใช้ได้');
ok(say('จำปี ตารางวันนี้')[0].text.indexOf('ตาราง') === 0, 'สั่งงานต่อท้ายชื่อบอทได้');
ok(!!say('จำปี ตารางวันนี้')[0].quickReply, 'คำตอบมีปุ่มลัดให้กดต่อ');
ok(say('วันนี้กินอะไรดี').length === 0, 'ข้อความอื่นในกลุ่ม บอทเงียบ');
ok(say('อะไรก็ได้', 'user').length === 1, 'แชทส่วนตัวตอบเสมอ');

console.log('\n' + (fails ? 'มีข้อทดสอบไม่ผ่าน ' + fails + ' ข้อ' : 'ผ่านทั้งหมด'));
process.exit(fails ? 1 : 0);
