/**
 * ระบบแจ้งตารางทำงานผ่าน LINE + อนุมัติโดย HR
 * วางไฟล์นี้ใน Google Sheets > ส่วนขยาย > Apps Script
 *
 * โครงสร้างข้อมูล 3 ชั้น
 *   พนักงาน    ทะเบียนคน ไม่เคยลบใคร คนลาออกใส่ "วันที่ออก"
 *   บันทึกเวร  แหล่งข้อมูลจริง 1 แถว = 1 คน 1 วัน ต่อท้ายลงล่างไปเรื่อย ๆ
 *   รอบ xxxx   ตารางแนวกว้างสำหรับดู/พิมพ์ สร้างใหม่จากบันทึกเวรได้ตลอด ลบทิ้งได้
 *
 * Script properties ที่ต้องตั้ง (Project Settings > Script properties):
 *   CHANNEL_ACCESS_TOKEN  = Channel access token (long-lived) ของ Messaging API
 *   LOGIN_CHANNEL_ID      = Channel ID ของ LINE Login channel (ที่มี LIFF)
 *   LIFF_ID               = LIFF ID เช่น 1234567890-AbCdEfGh
 */

const CFG = {
  BOT_NAME: 'จำปี',                                  // ชื่อที่ใช้เรียกบอทในกลุ่ม
  BOT_ALIASES: ['จำปี', 'จําปี', 'jampee', 'แจ้งงาน', 'เมนู'],  // คำขึ้นต้นที่ถือว่าเรียกบอท

  EMP_SHEET: 'พนักงาน',
  LOG_SHEET: 'บันทึกเวร',
  REQ_SHEET: 'คำขอ',
  USER_SHEET: 'ผู้ใช้ระบบ',
  VIEW_PREFIX: 'รอบ ',          // ชีทตารางแนวกว้าง เช่น "รอบ 2569-09"
  LEGACY_SHEET: 'รหัสพนักงาน',   // ชีทเดิมแนวกว้าง ใช้ตอนย้ายข้อมูลครั้งเดียว

  CUT_DAY: 26,                  // รอบสรุป = วันที่ 26 เดือนก่อน ถึง 25 เดือนนี้
  BACK_DAYS: 7,                 // ฟอร์มให้แจ้งย้อนหลังได้กี่วัน
  AHEAD_PERIODS: 1,             // ฟอร์มให้แจ้งล่วงหน้าอีกกี่รอบ

  SHIFTS: ['8.00-17.00', '10.00-18.00', '14.00-22.00', '18.00-02.00',
           '19.00-03.00', '20.00-04.00', '21.00-05.00'],
  STATUSES: ['หยุด', 'ลากิจ', 'ลาป่วย', 'ขาดงาน'],

  ROLE_LEAD: 'หัวหน้างาน',
  ROLE_HR: 'HR',

  ST_PENDING: 'รออนุมัติ',
  ST_OK: 'อนุมัติ',
  ST_NO: 'ไม่อนุมัติ',
  ST_OLD: 'แทนที่',             // ถูกคำขอที่อนุมัติทีหลังเขียนทับ
  PENDING: '(รอ) '
};

const EMP_HEADERS = ['รหัส', 'ชื่อ - สกุล', 'ชื่อเล่น', 'แผนก', 'ชื่อในกลุ่มไลน์',
  'ค่าแรง', 'หน่วย', 'วันเริ่มงาน', 'วันที่ออก', 'หมายเหตุ'];
const EMP_COL = { code: 1, name: 2, nick: 3, dept: 4, lineName: 5, wage: 6, unit: 7, from: 8, to: 9 };

const LOG_HEADERS = ['วันที่', 'รหัส', 'ชื่อ', 'กะ', 'สถานะ', 'เลขคำขอ', 'ผู้แจ้ง', 'ผู้พิจารณา', 'เวลาอัปเดต'];
const LOG_COL = { ymd: 1, code: 2, label: 3, shift: 4, status: 5, reqId: 6, by: 7, approver: 8, at: 9 };

const REQ_HEADERS = ['เลขคำขอ', 'เวลาส่ง', 'userId ผู้แจ้ง', 'ผู้แจ้ง', 'วันที่ทำงาน',
  'สรุป', 'รายการ (JSON)', 'คำเตือน (JSON)', 'สถานะ', 'ผู้พิจารณา', 'เวลาพิจารณา', 'เหตุผล/หมายเหตุ'];

const VIEW_HEAD = ['รหัส', 'ชื่อ - สกุล', 'ชื่อในกลุ่มไลน์', 'แผนก', 'ชื่อเล่น', 'ค่าแรง', 'หน่วย'];

/* ================= เมนูในชีท ================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('ตารางทำงาน')
    .addItem('สร้าง/รีเฟรชตารางรอบปัจจุบัน', 'buildCurrentPeriod')
    .addItem('สร้าง/รีเฟรชตารางรอบก่อนหน้า', 'buildPreviousPeriod')
    .addSeparator()
    .addItem('ตรวจการตั้งค่า', 'checkSetup')
    .addItem('ย้ายข้อมูลจากชีทเดิม (ครั้งเดียว)', 'migrateFromLegacy')
    .addToUi();
}

/* ================= ตั้งค่าครั้งแรก ================= */

/** กด Run ฟังก์ชันนี้ 1 ครั้งหลังวางโค้ด */
function setup() {
  const ss = ss_();

  if (!ss.getSheetByName(CFG.EMP_SHEET)) {
    const sh = ss.insertSheet(CFG.EMP_SHEET);
    sh.getRange(1, 1, 1, EMP_HEADERS.length).setValues([EMP_HEADERS]).setFontWeight('bold');
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('H:I').setNumberFormat('yyyy-mm-dd');
    sh.setFrozenRows(1);
    sh.setColumnWidth(2, 180);
  }

  if (!ss.getSheetByName(CFG.LOG_SHEET)) {
    const sh = ss.insertSheet(CFG.LOG_SHEET);
    sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]).setFontWeight('bold');
    sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
    sh.getRange('B:B').setNumberFormat('@');
    sh.getRange('I:I').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.setFrozenRows(1);
  }

  if (!ss.getSheetByName(CFG.REQ_SHEET)) {
    const sh = ss.insertSheet(CFG.REQ_SHEET);
    sh.getRange(1, 1, 1, REQ_HEADERS.length).setValues([REQ_HEADERS]).setFontWeight('bold');
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('E:E').setNumberFormat('@');
    sh.setFrozenRows(1);
  }

  if (!ss.getSheetByName(CFG.USER_SHEET)) {
    const sh = ss.insertSheet(CFG.USER_SHEET);
    sh.getRange(1, 1, 1, 4).setValues([['userId', 'ชื่อ', 'บทบาท (หัวหน้างาน / HR)',
      'แผนกที่ดูแล (คั่นด้วย , ว่าง = ทุกแผนก)']]).setFontWeight('bold');
    sh.getRange('C2:C200').setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList([CFG.ROLE_LEAD, CFG.ROLE_HR], true).build());
    sh.setFrozenRows(1);
  }

  checkSetup();
}

/** ตรวจว่าอ่านชีทได้ถูกต้อง ดูผลใน Execution log */
function checkSetup() {
  const log = [];
  const say = m => { log.push(m); console.log(m); };
  const warn = m => { log.push('! ' + m); console.warn(m); };

  say('เขตเวลาชีท: ' + tz_() + (tz_() === 'Asia/Bangkok' ? ' (ถูกต้อง)' : ' <-- ควรเป็น Asia/Bangkok'));

  [CFG.EMP_SHEET, CFG.LOG_SHEET, CFG.REQ_SHEET, CFG.USER_SHEET].forEach(n => {
    if (!ss_().getSheetByName(n)) warn('ยังไม่มีชีท "' + n + '" (กด Run ฟังก์ชัน setup)');
  });
  if (!ss_().getSheetByName(CFG.EMP_SHEET)) return log.join('\n');

  const today = ymd_(new Date());
  const p = period_(today);
  say('วันนี้ ' + today + ' อยู่ในรอบ ' + p.key + ' (' + p.label + ')');

  const emps = readEmployees_();
  const active = emps.filter(e => activeOn_(e, today));
  say('พนักงานในทะเบียน: ' + emps.length + ' คน  ทำงานอยู่วันนี้: ' + active.length + ' คน');
  active.forEach(e => say('  ' + e.code + '  ' + e.label + '  [' + e.dept + ']'));

  const left = emps.filter(e => e.to && e.to < today);
  if (left.length) say('ลาออกแล้ว (ไม่แสดงในฟอร์ม แต่ประวัติยังอยู่): ' + left.map(e => e.label).join(', '));

  const dup = {};
  active.forEach(e => (dup[e.label] = (dup[e.label] || 0) + 1));
  Object.keys(dup).filter(k => dup[k] > 1).forEach(k => warn('ชื่อแสดงซ้ำกัน: ' + k + ' (ควรแก้ชื่อเล่นให้ไม่ซ้ำ)'));

  const badDate = emps.filter(e => (e.from && !/^\d{4}-\d{2}-\d{2}$/.test(e.from)) || (e.to && !/^\d{4}-\d{2}-\d{2}$/.test(e.to)));
  badDate.forEach(e => warn('วันเริ่มงาน/วันที่ออกของ ' + e.code + ' ไม่ใช่วันที่ที่อ่านได้ (ต้องเป็นวันที่จริง ไม่ใช่ข้อความ)'));

  const rows = readLog_();
  say('บันทึกเวรทั้งหมด: ' + rows.length + ' แถว');
  const pend = rows.filter(r => r.status === CFG.ST_PENDING).length;
  if (pend) say('มีเวรที่ยังรออนุมัติ: ' + pend + ' รายการ');

  ['CHANNEL_ACCESS_TOKEN', 'LOGIN_CHANNEL_ID', 'LIFF_ID'].forEach(k => {
    if (!prop_(k)) warn('ยังไม่ได้ตั้ง Script property: ' + k);
  });
  return log.join('\n');
}

/* ================= จุดรับข้อมูล ================= */

function doGet() {
  return ContentService.createTextOutput('KB schedule API is running');
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad request' }); }
  if (body && body.events) { // มาจาก LINE webhook
    handleLine_(body.events);
    return ContentService.createTextOutput('OK');
  }
  return json_(handleApi_(body || {})); // มาจากหน้าฟอร์ม LIFF
}

/* ================= API สำหรับหน้าฟอร์ม ================= */

function handleApi_(b) {
  try {
    const who = verifyIdToken_(b.idToken);
    const user = getUser_(who.userId);
    switch (b.action) {
      case 'init': return apiInit_(who, user);
      case 'submit': return apiSubmit_(needRole_(user, [CFG.ROLE_LEAD, CFG.ROLE_HR]), b);
      case 'pending': needRole_(user, [CFG.ROLE_HR]); return { ok: true, pending: listPending_() };
      case 'approve': return apiApprove_(needRole_(user, [CFG.ROLE_HR]), b);
      case 'reject': return apiReject_(needRole_(user, [CFG.ROLE_HR]), b);
      default: throw new Error('ไม่รู้จักคำสั่ง ' + b.action);
    }
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/** ช่วงวันที่ที่ฟอร์มให้เลือกได้: ย้อนหลัง BACK_DAYS วัน ถึงสิ้นรอบถัดไป */
function formRange_() {
  const today = ymd_(new Date());
  let p = period_(today);
  for (let i = 0; i < CFG.AHEAD_PERIODS; i++) p = nextPeriod_(p);
  return { from: addDays_(today, -CFG.BACK_DAYS), to: p.end };
}

function apiInit_(who, user) {
  const out = { ok: true, me: who, user: user };
  if (!user) return out;

  const rng = formRange_();
  const today = ymd_(new Date());
  const dates = datesBetween_(rng.from, rng.to);
  // แสดงเฉพาะคนที่ยังทำงานอยู่ตั้งแต่วันนี้เป็นต้นไป คนลาออกแล้วจะหายจากฟอร์ม
  const emps = readEmployees_().filter(e => activeBetween_(e, today, rng.to));

  out.employees = emps.map(e => ({ code: e.code, label: e.label, name: e.name, dept: e.dept }));
  out.dates = dates.map(d => ({ ymd: d, label: thDate_(d, true) }));
  out.shifts = CFG.SHIFTS;
  out.statuses = CFG.STATUSES;

  const tomorrow = addDays_(today, 1);
  out.defaultDate = dates.indexOf(tomorrow) >= 0 ? tomorrow : (dates.indexOf(today) >= 0 ? today : dates[0] || null);

  if (user.role === CFG.ROLE_HR) {
    out.pending = listPending_();
    out.noCode = readEmployeeWarnings_();
  }
  return out;
}

function apiSubmit_(user, p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const rng = formRange_();
    if (!p.date || p.date < rng.from || p.date > rng.to) {
      throw new Error('วันที่ ' + p.date + ' อยู่นอกช่วงที่แจ้งได้ (' + rng.from + ' ถึง ' + rng.to + ')');
    }

    const byCode = indexBy_(readEmployees_().filter(e => activeOn_(e, p.date)));
    const seen = {};
    const entries = [];
    (p.groups || []).forEach(g => {
      const shift = normShift_(g.shift);
      (g.codes || []).forEach(c => {
        const e = byCode[c];
        if (!e) throw new Error('ไม่พบรหัสพนักงาน ' + c + ' ที่ทำงานอยู่ในวันที่ ' + p.date);
        if (seen[c]) throw new Error(e.label + ' ถูกเลือกมากกว่า 1 กะ');
        seen[c] = 1;
        entries.push({ code: c, label: e.label, shift: shift });
      });
    });
    if (!entries.length) throw new Error('ยังไม่ได้เลือกพนักงาน');

    const state = dayState_(readLog_(), p.date);
    const warnings = [];
    entries.forEach(en => {
      const s = state[en.code] || {};
      en.old = s.ok ? s.ok.shift : '';
      if (s.pending) warnings.push(en.label + ': มีคำขออื่นรออยู่ (#' + s.pending.reqId + ') คำขอนี้จะแทนที่เมื่ออนุมัติ');
    });

    const id = nextId_();
    const now = new Date();
    appendLog_(entries.map(en => {
      const row = [];
      row[LOG_COL.ymd - 1] = p.date;
      row[LOG_COL.code - 1] = en.code;
      row[LOG_COL.label - 1] = en.label;
      row[LOG_COL.shift - 1] = en.shift;
      row[LOG_COL.status - 1] = CFG.ST_PENDING;
      row[LOG_COL.reqId - 1] = id;
      row[LOG_COL.by - 1] = user.name;
      row[LOG_COL.approver - 1] = '';
      row[LOG_COL.at - 1] = now;
      return row;
    }));

    ss_().getSheetByName(CFG.REQ_SHEET).appendRow([
      id, now, user.userId, user.name, p.date, summaryText_(entries),
      JSON.stringify(entries), JSON.stringify(warnings), CFG.ST_PENDING, '', '', ''
    ]);
    SpreadsheetApp.flush();

    const req = { id: id, fromName: user.name, date: p.date, entries: entries, warnings: warnings };
    try {
      hrUsers_().forEach(h => push_(h.userId, [hrCard_(req)]));
    } catch (err) {
      console.error('แจ้ง HR ไม่สำเร็จ', err);
    }
    return { ok: true, id: id, dateLabel: thDate_(p.date, true), groups: groupByShift_(entries), warnings: warnings };
  } finally {
    lock.releaseLock();
  }
}

function apiApprove_(user, p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = findReq_(p.id);
    if (req.status !== CFG.ST_PENDING) throw new Error('คำขอ #' + req.id + ' ถูกพิจารณาไปแล้ว (' + req.status + ')');

    const edits = p.edits || {};
    const changed = [];
    req.entries.forEach(en => {
      if (edits[en.code] && normShift_(edits[en.code]) !== en.shift) {
        changed.push(en.label + ' ' + en.shift + ' → ' + normShift_(edits[en.code]));
        en.shift = normShift_(edits[en.code]);
      }
    });

    const log = readLog_();
    const mine = {};
    log.forEach(r => { if (r.reqId === req.id && r.status === CFG.ST_PENDING) mine[r.code] = r; });

    const skipped = [];
    const now = new Date();
    const ok = [];
    req.entries.forEach(en => {
      const row = mine[en.code];
      if (!row) { skipped.push(en.label); return; }
      // เวรที่อนุมัติไว้ก่อนหน้าของคน-วันเดียวกัน ถือว่าถูกแทนที่
      log.forEach(r => {
        if (r.row !== row.row && r.ymd === row.ymd && r.code === row.code && r.status === CFG.ST_OK) {
          setLogStatus_(r.row, CFG.ST_OLD, user.name, now);
        }
      });
      setLogShift_(row.row, en.shift);
      setLogStatus_(row.row, CFG.ST_OK, user.name, now);
      ok.push(en.label);
    });

    const note = [changed.length ? 'แก้ไข: ' + changed.join(', ') : '',
                  skipped.length ? 'ข้าม (ไม่พบรายการรออนุมัติ): ' + skipped.join(', ') : ''].filter(String).join(' | ');
    updateReq_(req, CFG.ST_OK, user.name, note);
    SpreadsheetApp.flush();

    let msg = 'คำขอ #' + req.id + ' ตาราง ' + thDate_(req.date, true) + ' อนุมัติแล้ว\n' + summaryText_(req.entries);
    if (changed.length) msg += '\n\nHR แก้ไข: ' + changed.join(', ');
    if (skipped.length) msg += '\n\nไม่ได้ลงตาราง: ' + skipped.join(', ');
    safePush_(req.fromUserId, [text_(msg)]);
    return { ok: true, id: req.id, skipped: skipped };
  } finally {
    lock.releaseLock();
  }
}

function apiReject_(user, p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = findReq_(p.id);
    if (req.status !== CFG.ST_PENDING) throw new Error('คำขอ #' + req.id + ' ถูกพิจารณาไปแล้ว (' + req.status + ')');

    const now = new Date();
    readLog_().forEach(r => {
      if (r.reqId === req.id && r.status === CFG.ST_PENDING) setLogStatus_(r.row, CFG.ST_NO, user.name, now);
    });

    const reason = String(p.reason || '').trim();
    updateReq_(req, CFG.ST_NO, user.name, reason);
    SpreadsheetApp.flush();

    safePush_(req.fromUserId, [text_('คำขอ #' + req.id + ' ตาราง ' + thDate_(req.date, true) +
      ' ไม่อนุมัติ' + (reason ? '\nเหตุผล: ' + reason : '') + '\n\nตารางยังเป็นค่าเดิม')]);
    return { ok: true, id: req.id };
  } finally {
    lock.releaseLock();
  }
}

/* ================= LINE webhook ================= */

function handleLine_(events) {
  events.forEach(ev => {
    try {
      if (ev.type === 'join' || ev.type === 'memberJoined') {
        reply_(ev.replyToken, [text_('สวัสดีครับ ผมชื่อ ' + CFG.BOT_NAME +
          ' ดูแลเรื่องตารางทำงาน\nพิมพ์ "' + CFG.BOT_NAME + '" เมื่อไหร่ก็ได้ เดี๋ยวผมส่งปุ่มให้กดครับ'), menuFlex_()]);
        return;
      }
      if (ev.type !== 'message' || ev.message.type !== 'text') return;

      const raw = ev.message.text.trim();
      const cmd = botCommand_(raw);
      const isDirect = ev.source && ev.source.type === 'user';

      if (cmd === null && !isDirect) return;   // ข้อความอื่นในกลุ่ม บอทจะเงียบ

      const t = (cmd === null ? raw : cmd).trim();
      if (!t || t === 'เมนู' || t === 'แจ้งงาน') reply_(ev.replyToken, [menuFlex_()]);
      else if (t === 'ดูตารางวันนี้' || t === 'ตารางวันนี้' || t === 'วันนี้') replyQ_(ev.replyToken, scheduleText_(0));
      else if (t === 'ดูตารางพรุ่งนี้' || t === 'ตารางพรุ่งนี้' || t === 'พรุ่งนี้') replyQ_(ev.replyToken, scheduleText_(1));
      else if (t.toLowerCase() === 'myid') {
        replyQ_(ev.replyToken, 'userId ของคุณ:\n' + (ev.source.userId || '(ไม่พบ ลองพิมพ์ในแชทส่วนตัวกับบอท)'));
      }
      else reply_(ev.replyToken, [menuFlex_()]);   // เรียกชื่อแล้วสั่งอะไรไม่รู้จัก ส่งเมนูให้กด
    } catch (err) {
      console.error(err);
    }
  });
}

/**
 * ตัดชื่อบอทออกจากต้นข้อความ
 * คืนคำสั่งที่เหลือ (อาจเป็นสตริงว่าง = เรียกชื่อเฉย ๆ) หรือ null ถ้าไม่ได้เรียกบอท
 */
function botCommand_(raw) {
  const s = raw.replace(/^@/, '').trim();
  const low = s.toLowerCase();
  for (let i = 0; i < CFG.BOT_ALIASES.length; i++) {
    const a = CFG.BOT_ALIASES[i].toLowerCase();
    if (low === a) return '';
    if (low.indexOf(a) === 0) return s.slice(a.length).replace(/^[\s,:：]+/, '');
  }
  return null;
}

function scheduleText_(offsetDays) {
  const ymd = addDays_(ymd_(new Date()), offsetDays);
  const state = dayState_(readLog_(), ymd);
  const groups = {};
  Object.keys(state).forEach(code => {
    const s = state[code];
    const v = s.pending ? CFG.PENDING + s.pending.shift : (s.ok ? s.ok.shift : '');
    if (!v) return;
    const label = (s.pending || s.ok).label || code;
    (groups[v] = groups[v] || []).push(label);
  });
  const keys = Object.keys(groups).sort((a, b) => shiftSortKey_(a) - shiftSortKey_(b));
  if (!keys.length) return 'ตาราง ' + thDate_(ymd, true) + '\nยังไม่มีข้อมูล';
  return 'ตาราง ' + thDate_(ymd, true) + '\n\n' + keys.map(k => k + '\n' + groups[k].join(', ')).join('\n\n');
}

/* ================= ข้อความ LINE ================= */

function liffUrl_(query) {
  return 'https://liff.line.me/' + prop_('LIFF_ID') + (query ? '?' + query : '');
}

function menuFlex_() {
  const btn = (label, action, primary) => ({
    type: 'button', style: primary ? 'primary' : 'secondary', color: primary ? '#2F6B45' : undefined,
    height: 'sm', margin: 'sm', action: Object.assign({ label: label }, action)
  });
  return {
    type: 'flex', altText: 'เมนูตารางทำงาน', quickReply: quickReply_(),
    contents: {
      type: 'bubble', size: 'kilo',
      body: {
        type: 'box', layout: 'vertical', contents: [
          { type: 'text', text: CFG.BOT_NAME + ' · ตารางทำงาน', weight: 'bold', size: 'lg' },
          { type: 'text', text: 'กดปุ่มได้เลย ไม่ต้องพิมพ์', size: 'sm', color: '#5E6B61', margin: 'xs' },
          btn('แจ้งวันทำงาน', { type: 'uri', uri: liffUrl_() }, true),
          btn('รายการรออนุมัติ (HR)', { type: 'uri', uri: liffUrl_('page=approve') }),
          btn('ดูตารางวันนี้', { type: 'message', text: CFG.BOT_NAME + ' ตารางวันนี้' }),
          btn('ดูตารางพรุ่งนี้', { type: 'message', text: CFG.BOT_NAME + ' ตารางพรุ่งนี้' })
        ]
      }
    }
  };
}

function hrCard_(req) {
  const rows = [];
  groupByShift_(req.entries).forEach(g => {
    rows.push({ type: 'text', text: g.shift, weight: 'bold', size: 'sm', margin: 'md' });
    rows.push({ type: 'text', text: g.names.join(', '), size: 'sm', wrap: true, color: '#1D2B21' });
  });
  (req.warnings || []).forEach(w => rows.push({ type: 'text', text: '⚠ ' + w, size: 'xs', wrap: true, color: '#B8741A', margin: 'sm' }));
  return {
    type: 'flex', altText: 'คำขอ #' + req.id + ' รออนุมัติ',
    contents: {
      type: 'bubble', size: 'kilo',
      body: {
        type: 'box', layout: 'vertical', contents: [
          { type: 'box', layout: 'horizontal', contents: [
            { type: 'text', text: 'คำขอ #' + req.id, weight: 'bold', flex: 1 },
            { type: 'text', text: 'รออนุมัติ', size: 'xs', color: '#B8741A', align: 'end', gravity: 'center' }
          ] },
          { type: 'text', text: 'จาก ' + req.fromName, size: 'xs', color: '#5E6B61' },
          { type: 'text', text: 'วันที่ ' + thDate_(req.date, true) + ' · ' + req.entries.length + ' คน', size: 'sm', margin: 'md' }
        ].concat(rows)
      },
      footer: {
        type: 'box', layout: 'vertical', contents: [
          { type: 'button', style: 'primary', color: '#2F6B45', height: 'sm',
            action: { type: 'uri', label: 'ตรวจและอนุมัติ', uri: liffUrl_('page=approve&id=' + req.id) } }
        ]
      }
    }
  };
}

function text_(t) { return { type: 'text', text: String(t).slice(0, 4900) }; }

/** ปุ่มลัดใต้ช่องพิมพ์ กดต่อได้เลยโดยไม่ต้องพิมพ์ชื่อบอทซ้ำ */
function quickReply_() {
  const item = (label, action) => ({ type: 'action', action: Object.assign({ label: label }, action) });
  return { items: [
    item('แจ้งวันทำงาน', { type: 'uri', uri: liffUrl_() }),
    item('ตารางวันนี้', { type: 'message', text: CFG.BOT_NAME + ' ตารางวันนี้' }),
    item('ตารางพรุ่งนี้', { type: 'message', text: CFG.BOT_NAME + ' ตารางพรุ่งนี้' }),
    item('เมนู', { type: 'message', text: CFG.BOT_NAME })
  ] };
}

/** ตอบข้อความพร้อมปุ่มลัด */
function replyQ_(token, t) {
  const m = text_(t);
  m.quickReply = quickReply_();
  return reply_(token, [m]);
}

function reply_(token, messages) { return lineApi_('message/reply', { replyToken: token, messages: messages }); }
function push_(to, messages) { return lineApi_('message/push', { to: to, messages: messages }); }
function safePush_(to, messages) { try { if (to) push_(to, messages); } catch (e) { console.error(e); } }

function lineApi_(path, payload) {
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/' + path, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + prop_('CHANNEL_ACCESS_TOKEN') },
    payload: JSON.stringify(payload)
  });
  if (res.getResponseCode() >= 300) console.error('LINE ' + path + ' ' + res.getResponseCode() + ': ' + res.getContentText());
  return res;
}

/** ตรวจ ID token จาก LIFF ว่าเป็นผู้ใช้จริง */
function verifyIdToken_(token) {
  if (!token) throw new Error('ไม่พบข้อมูลยืนยันตัวตน กรุณาเปิดฟอร์มจากในแอป LINE');
  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post', muteHttpExceptions: true,
    payload: { id_token: token, client_id: prop_('LOGIN_CHANNEL_ID') }
  });
  if (res.getResponseCode() !== 200) throw new Error('ยืนยันตัวตนไม่สำเร็จ กรุณาปิดแล้วเปิดฟอร์มใหม่');
  const d = JSON.parse(res.getContentText());
  return { userId: d.sub, name: d.name || '' };
}

/* ================= ชีทพนักงาน ================= */

function readEmployees_() {
  const sh = need_(CFG.EMP_SHEET);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const v = sh.getRange(2, 1, last - 1, EMP_HEADERS.length).getValues();
  const out = [];
  v.forEach((r, i) => {
    const code = String(r[EMP_COL.code - 1]).trim();
    if (!code) return;
    const name = String(r[EMP_COL.name - 1]).trim();
    const nick = String(r[EMP_COL.nick - 1]).trim();
    out.push({
      row: i + 2, code: code, name: name, nick: nick,
      label: nick || shortName_(name) || code,
      dept: String(r[EMP_COL.dept - 1]).trim() || 'ไม่ระบุแผนก',
      lineName: String(r[EMP_COL.lineName - 1]).trim(),
      wage: r[EMP_COL.wage - 1], unit: String(r[EMP_COL.unit - 1]).trim(),
      from: toYmd_(r[EMP_COL.from - 1]), to: toYmd_(r[EMP_COL.to - 1])
    });
  });
  return out;
}

/** คนที่กรอกชื่อไว้แต่ยังไม่มีรหัส จะไม่แสดงในฟอร์ม */
function readEmployeeWarnings_() {
  const sh = need_(CFG.EMP_SHEET);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 2).getValues()
    .filter(r => !String(r[0]).trim() && String(r[1]).trim())
    .map(r => String(r[1]).trim());
}

function activeOn_(e, ymd) {
  return (!e.from || e.from <= ymd) && (!e.to || e.to >= ymd);
}
function activeBetween_(e, a, b) {
  return (!e.from || e.from <= b) && (!e.to || e.to >= a);
}

/* ================= ชีทบันทึกเวร ================= */

function readLog_() {
  const sh = need_(CFG.LOG_SHEET);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const v = sh.getRange(2, 1, last - 1, LOG_HEADERS.length).getValues();
  const out = [];
  v.forEach((r, i) => {
    const code = String(r[LOG_COL.code - 1]).trim();
    const ymd = toYmd_(r[LOG_COL.ymd - 1]);
    if (!code || !ymd) return;
    out.push({
      row: i + 2, ymd: ymd, code: code, label: String(r[LOG_COL.label - 1]).trim(),
      shift: String(r[LOG_COL.shift - 1]).trim(), status: String(r[LOG_COL.status - 1]).trim(),
      reqId: String(r[LOG_COL.reqId - 1]).trim(), by: String(r[LOG_COL.by - 1]).trim(),
      approver: String(r[LOG_COL.approver - 1]).trim(), at: r[LOG_COL.at - 1]
    });
  });
  return out;
}

/** สถานะล่าสุดของแต่ละคนในวันหนึ่ง { code: { ok: row, pending: row } } */
function dayState_(log, ymd) {
  const m = {};
  log.forEach(r => {
    if (r.ymd !== ymd) return;
    const s = m[r.code] = m[r.code] || {};
    if (r.status === CFG.ST_OK) s.ok = r;
    else if (r.status === CFG.ST_PENDING) s.pending = r;
  });
  return m;
}

function appendLog_(rows) {
  if (!rows.length) return [];
  const sh = need_(CFG.LOG_SHEET);
  const start = Math.max(sh.getLastRow(), 1) + 1;
  sh.getRange(start, 1, rows.length, LOG_HEADERS.length).setValues(rows);
  return rows.map((_, i) => start + i);
}

function setLogShift_(row, shift) {
  need_(CFG.LOG_SHEET).getRange(row, LOG_COL.shift).setValue(shift);
}

function setLogStatus_(row, status, approver, at) {
  const sh = need_(CFG.LOG_SHEET);
  sh.getRange(row, LOG_COL.status).setValue(status);
  sh.getRange(row, LOG_COL.approver, 1, 2).setValues([[approver || '', at || new Date()]]);
}

/* ================= ชีทคำขอ / ผู้ใช้ระบบ ================= */

function getUser_(userId) {
  const sh = ss_().getSheetByName(CFG.USER_SHEET);
  if (!sh) return null;
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === userId) {
      return {
        userId: userId, name: String(v[i][1]).trim() || 'ไม่ระบุชื่อ', role: String(v[i][2]).trim(),
        depts: String(v[i][3]).split(',').map(x => x.trim()).filter(String)
      };
    }
  }
  return null;
}

function hrUsers_() {
  const sh = ss_().getSheetByName(CFG.USER_SHEET);
  return sh.getDataRange().getValues().slice(1)
    .filter(r => String(r[2]).trim() === CFG.ROLE_HR && String(r[0]).trim())
    .map(r => ({ userId: String(r[0]).trim(), name: String(r[1]) }));
}

function needRole_(user, roles) {
  if (!user) throw new Error('บัญชีนี้ยังไม่มีสิทธิ์ใช้งาน');
  if (roles.indexOf(user.role) < 0) throw new Error('เมนูนี้สำหรับ ' + roles.join(' / ') + ' เท่านั้น');
  return user;
}

function nextId_() {
  const sh = need_(CFG.REQ_SHEET);
  return 'R' + ('000' + sh.getLastRow()).slice(-4);
}

function readReqRows_() {
  const sh = need_(CFG.REQ_SHEET);
  const v = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    if (!r[0]) continue;
    out.push({
      row: i + 1, id: String(r[0]), createdAt: r[1], fromUserId: String(r[2]), fromName: String(r[3]),
      date: toYmd_(r[4]), summary: String(r[5]),
      entries: safeJson_(r[6], []), warnings: safeJson_(r[7], []), status: String(r[8])
    });
  }
  return out;
}

function findReq_(id) {
  const r = readReqRows_().find(x => x.id === id);
  if (!r) throw new Error('ไม่พบคำขอ #' + id);
  return r;
}

function updateReq_(req, status, by, note) {
  need_(CFG.REQ_SHEET).getRange(req.row, 6, 1, 7).setValues([[summaryText_(req.entries),
    JSON.stringify(req.entries), JSON.stringify(req.warnings), status, by, new Date(), note || '']]);
}

function listPending_() {
  return readReqRows_().filter(r => r.status === CFG.ST_PENDING).map(r => ({
    id: r.id, fromName: r.fromName, date: r.date, dateLabel: thDate_(r.date, true),
    createdLabel: r.createdAt instanceof Date ? Utilities.formatDate(r.createdAt, tz_(), 'd/M HH:mm') : '',
    entries: r.entries.map(e => ({ code: e.code, label: e.label, shift: e.shift, old: e.old || '' })),
    warnings: r.warnings
  }));
}

/* ================= ตารางแนวกว้างของแต่ละรอบ ================= */

function buildCurrentPeriod() { buildPeriodSheet_(period_(ymd_(new Date()))); }
function buildPreviousPeriod() { buildPeriodSheet_(prevPeriod_(period_(ymd_(new Date())))); }

/** สร้าง/รีเฟรชชีท "รอบ xxxx" จากบันทึกเวร ลบทิ้งแล้วสร้างใหม่ได้เสมอ */
function buildPeriodSheet_(p) {
  const ss = ss_();
  const dates = periodDates_(p);
  const emps = readEmployees_().filter(e => activeBetween_(e, p.start, p.end));
  const log = readLog_();
  const state = {};
  dates.forEach(d => (state[d] = dayState_(log, d)));

  const name = CFG.VIEW_PREFIX + p.key;
  let sh = ss.getSheetByName(name);
  if (sh) { sh.clear(); sh.clearConditionalFormatRules(); sh.getDataRange().clearNote(); }
  else sh = ss.insertSheet(name);

  const nCol = VIEW_HEAD.length + dates.length;

  // แถว 1 หัวเรื่อง
  sh.getRange(1, 1).setValue('ตารางทำงาน วันที่ ' + p.label);
  sh.getRange(1, 1, 1, nCol).merge().setFontWeight('bold').setHorizontalAlignment('center');

  // แถว 2 หัวตาราง + วันที่
  sh.getRange(2, 1, 1, VIEW_HEAD.length).setValues([VIEW_HEAD]);
  const dateCells = dates.map(d => { const q = ymdParts_(d); return new Date(q[0], q[1] - 1, q[2]); });
  const dateRange = sh.getRange(2, VIEW_HEAD.length + 1, 1, dates.length);
  dateRange.setValues([dateCells]).setNumberFormat('d-mmm');
  sh.getRange(2, 1, 1, nCol).setFontWeight('bold').setHorizontalAlignment('center');

  // คอลัมน์วันอาทิตย์ = พื้นเหลือง (ของเดิมใช้สีนี้แทนวันหยุด)
  dates.forEach((d, i) => {
    if (dateCells[i].getDay() === 0) sh.getRange(2, VIEW_HEAD.length + 1 + i).setBackground('#FFF2B2');
  });

  // แถว 4 เป็นต้นไป
  const body = emps.map(e => {
    const head = [e.code, e.name, e.lineName, e.dept, e.nick, e.wage, e.unit];
    const cells = dates.map(d => {
      const s = (state[d] || {})[e.code] || {};
      if (s.pending) return CFG.PENDING + s.pending.shift;
      return s.ok ? s.ok.shift : '';
    });
    return head.concat(cells);
  });
  if (body.length) {
    sh.getRange(4, 1, body.length, nCol).setValues(body);
    sh.getRange(4, 6, body.length, 1).setNumberFormat('#,##0.00'); // ค่าแรง
  }

  // โน้ตเลขคำขอบนช่องที่ยังรออนุมัติ
  emps.forEach((e, r) => dates.forEach((d, c) => {
    const s = (state[d] || {})[e.code] || {};
    if (s.pending) sh.getRange(4 + r, VIEW_HEAD.length + 1 + c).setNote('รออนุมัติ #' + s.pending.reqId + ' โดย ' + s.pending.by);
  }));

  // สีตามความหมาย ใช้การจัดรูปแบบตามเงื่อนไข ไม่ใช่สีที่ระบายมือ
  if (body.length) {
    const area = sh.getRange(4, VIEW_HEAD.length + 1, body.length, dates.length);
    const rules = [
      SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith(CFG.PENDING)
        .setBackground('#FFE4C4').setFontColor('#8A4B08').setRanges([area]).build()
    ];
    CFG.STATUSES.forEach(s => rules.push(
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s)
        .setBackground('#FFF2B2').setRanges([area]).build()));
    sh.setConditionalFormatRules(rules);
  }

  sh.setFrozenRows(3);
  sh.setFrozenColumns(5);
  sh.setColumnWidth(2, 180);
  SpreadsheetApp.flush();
  return sh.getName();
}

/* ================= ย้ายข้อมูลจากชีทเดิม (ครั้งเดียว) ================= */

/**
 * อ่านชีท "รหัสพนักงาน" แบบกว้าง แล้วเติมลงชีท พนักงาน + บันทึกเวร
 * รันซ้ำได้ ระบบจะข้ามคนและวันที่ที่ย้ายไปแล้ว
 */
function migrateFromLegacy() {
  const ss = ss_();
  const sh = ss.getSheetByName(CFG.LEGACY_SHEET);
  if (!sh) throw new Error('ไม่พบชีท "' + CFG.LEGACY_SHEET + '" (ถ้าย้ายเสร็จแล้วไม่ต้องรันซ้ำ)');
  setup();

  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();

  // แถว 2 ตั้งแต่คอลัมน์ H: เก็บทุกช่องที่เป็นวันที่ ข้ามช่องว่างระหว่างบล็อกได้
  const head = sh.getRange(2, 8, 1, Math.max(lastCol - 7, 1)).getValues()[0];
  const cols = [];
  head.forEach((v, i) => { if (v instanceof Date) cols.push({ ymd: ymd_(v), col: 8 + i }); });
  if (!cols.length) throw new Error('ไม่พบวันที่ในแถว 2 ของชีท "' + CFG.LEGACY_SHEET + '"');

  const grid = sh.getRange(4, 1, Math.max(lastRow - 3, 1), lastCol).getDisplayValues();

  const existing = {};
  readEmployees_().forEach(e => (existing[e.code] = e));
  const done = {};
  readLog_().forEach(r => (done[r.ymd + '|' + r.code] = 1));

  const newEmps = [];
  const newLogs = [];
  const now = new Date();

  for (let i = 0; i < grid.length; i++) {
    const r = grid[i];
    const code = String(r[0]).trim();
    if (code === 'หมายเหตุ') break;
    if (!code) continue;

    if (!existing[code]) {
      const row = [];
      row[EMP_COL.code - 1] = code;
      row[EMP_COL.name - 1] = String(r[1]).trim();
      row[EMP_COL.nick - 1] = String(r[4]).trim();
      row[EMP_COL.dept - 1] = String(r[3]).trim();
      row[EMP_COL.lineName - 1] = String(r[2]).trim();
      row[EMP_COL.wage - 1] = String(r[5]).trim();
      row[EMP_COL.unit - 1] = String(r[6]).trim();
      row[EMP_COL.from - 1] = '';
      row[EMP_COL.to - 1] = '';
      row[EMP_HEADERS.length - 1] = 'ย้ายจากชีทเดิม';
      newEmps.push(row);
      existing[code] = { code: code, label: String(r[4]).trim() || shortName_(String(r[1])) || code };
    }
    const label = existing[code].label || code;

    cols.forEach(c => {
      const v = String(r[c.col - 1] || '').trim().replace(CFG.PENDING, '');
      if (!v || done[c.ymd + '|' + code]) return;
      const row = [];
      row[LOG_COL.ymd - 1] = c.ymd;
      row[LOG_COL.code - 1] = code;
      row[LOG_COL.label - 1] = label;
      row[LOG_COL.shift - 1] = v;
      row[LOG_COL.status - 1] = CFG.ST_OK;
      row[LOG_COL.reqId - 1] = '';
      row[LOG_COL.by - 1] = 'ย้ายจากชีทเดิม';
      row[LOG_COL.approver - 1] = '';
      row[LOG_COL.at - 1] = now;
      newLogs.push(row);
    });
  }

  if (newEmps.length) {
    const es = need_(CFG.EMP_SHEET);
    es.getRange(Math.max(es.getLastRow(), 1) + 1, 1, newEmps.length, EMP_HEADERS.length).setValues(newEmps);
  }
  appendLog_(newLogs);
  SpreadsheetApp.flush();

  const msg = 'ย้ายข้อมูลแล้ว: พนักงานใหม่ ' + newEmps.length + ' คน, บันทึกเวร ' + newLogs.length + ' แถว' +
    ' (ช่วง ' + cols[0].ymd + ' ถึง ' + cols[cols.length - 1].ymd + ')';
  console.log(msg);
  return msg;
}

/* ================= รอบวันที่ ================= */

/** รอบที่ครอบคลุมวันนี้: 26 เดือนก่อน ถึง 25 เดือนนี้ */
function period_(ymd) {
  const q = ymdParts_(ymd);
  let ey = q[0], em = q[1];
  if (q[2] >= CFG.CUT_DAY) { em++; if (em > 12) { em = 1; ey++; } }
  let sy = ey, sm = em - 1;
  if (sm < 1) { sm = 12; sy--; }
  return {
    key: (ey + 543) + '-' + p2_(em),
    start: mkYmd_(sy, sm, CFG.CUT_DAY),
    end: mkYmd_(ey, em, CFG.CUT_DAY - 1),
    label: CFG.CUT_DAY + ' ' + TH_MONTHS_FULL[sm - 1] + ' ' + (sy + 543) +
           ' - ' + (CFG.CUT_DAY - 1) + ' ' + TH_MONTHS_FULL[em - 1] + ' ' + (ey + 543)
  };
}
function nextPeriod_(p) { return period_(addDays_(p.end, 1)); }
function prevPeriod_(p) { return period_(addDays_(p.start, -1)); }
function periodDates_(p) { return datesBetween_(p.start, p.end); }

function datesBetween_(a, b) {
  const out = [];
  for (let d = a; d <= b; d = addDays_(d, 1)) out.push(d);
  return out;
}

/* ================= เครื่องมือย่อย ================= */

function normShift_(s) {
  s = String(s || '').trim().replace(/:/g, '.').replace(/\s+/g, '');
  if (CFG.STATUSES.indexOf(s) >= 0) return s;
  const m = s.match(/^(\d{1,2})\.(\d{2})-(\d{1,2})\.(\d{2})$/);
  if (!m || +m[1] > 23 || +m[3] > 23 || +m[2] > 59 || +m[4] > 59) {
    throw new Error('รูปแบบเวลาไม่ถูกต้อง "' + s + '" (ตัวอย่าง 21.00-05.00)');
  }
  return (+m[1]) + '.' + m[2] + '-' + ('0' + (+m[3])).slice(-2) + '.' + m[4]; // เช่น 8.00-17.00, 21.00-05.00
}

function shiftSortKey_(s) {
  const m = String(s).replace(CFG.PENDING, '').match(/^(\d{1,2})\.(\d{2})/);
  return m ? (+m[1]) * 60 + (+m[2]) : 10000;
}

function groupByShift_(entries) {
  const g = {};
  entries.forEach(e => (g[e.shift] = g[e.shift] || []).push(e.label));
  return Object.keys(g).sort((a, b) => shiftSortKey_(a) - shiftSortKey_(b)).map(k => ({ shift: k, names: g[k] }));
}

function summaryText_(entries) {
  return groupByShift_(entries).map(g => g.shift + ': ' + g.names.join(', ')).join('\n');
}

function shortName_(name) {
  return String(name).replace(/^(น\.ส\.|นางสาว|นาย|นาง)\s*/, '').split(/\s+/)[0];
}

const TH_MONTHS = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
const TH_MONTHS_FULL = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
const TH_DAYS = ['อา.', 'จ.', 'อ.', 'พ.', 'พฤ.', 'ศ.', 'ส.'];

function thDate_(ymd, withDay) {
  const p = ymdParts_(ymd);
  const d = new Date(p[0], p[1] - 1, p[2]);
  return (withDay ? TH_DAYS[d.getDay()] + ' ' : '') + p[2] + ' ' + TH_MONTHS[p[1] - 1] + ' ' + String(p[0] + 543).slice(-2);
}

function ymdParts_(s) { return String(s).split('-').map(Number); }
function p2_(n) { return ('0' + n).slice(-2); }
function mkYmd_(y, m, d) { return y + '-' + p2_(m) + '-' + p2_(d); }
function addDays_(ymd, n) {
  const q = ymdParts_(ymd);
  const t = new Date(q[0], q[1] - 1, q[2] + n);
  return mkYmd_(t.getFullYear(), t.getMonth() + 1, t.getDate());
}
function toYmd_(v) { return v instanceof Date ? ymd_(v) : String(v == null ? '' : v).trim(); }

function indexBy_(list) { const o = {}; list.forEach(e => (o[e.code] = e)); return o; }
function safeJson_(v, dflt) { try { return JSON.parse(v); } catch (e) { return dflt; } }
function need_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('ไม่พบชีท "' + name + '" (กด Run ฟังก์ชัน setup ก่อน)');
  return sh;
}
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function tz_() { return ss_().getSpreadsheetTimeZone(); }
function ymd_(d) { return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'); }
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
