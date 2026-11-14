/**
 * ระบบเช็คชื่อผู้เรียน — ส่วนหลังบ้าน (Google Apps Script + Google Sheet)
 *
 * สคริปต์นี้ต้อง "ผูกกับ Google Sheet" (เปิดชีต > ส่วนขยาย > Apps Script)
 * แล้ว Deploy เป็น Web app — ดูขั้นตอนเต็มใน README.md
 *
 * ชีตที่ระบบสร้างให้อัตโนมัติ:
 *   • เช็คชื่อ  — 1 แถวต่อการเช็คชื่อ 1 ครั้ง (วันที่ เวลา ชื่อ สถานะ ลายเซ็น)
 *   • รายชื่อ   — รายชื่อผู้เรียน (แก้ไขในชีตได้โดยตรง คอลัมน์ id ปล่อยว่างได้)
 *   • ตั้งค่า   — ชื่อคลาส เวลาเริ่มเรียน และจำนวนนาทีที่ถือว่าสาย
 */

// ======================= ตั้งค่า =======================
// รหัสผู้ดูแล: ใช้ปลดล็อกเมนู "จัดการ" บนเว็บ (นำเข้ารายชื่อ / ตั้งค่า / ยกเลิกการเช็คชื่อ)
// ⚠️ เปลี่ยนก่อนใช้งานจริง และอย่าใส่รหัสนี้ไว้ในไฟล์ที่ขึ้น GitHub
const ADMIN_KEY = 'change-me-1234';

// เก็บรูปลายเซ็นเป็นไฟล์ PNG ใน Google Drive ของคุณด้วย (นอกจากในชีต)
const SAVE_SIGNATURE_TO_DRIVE = true;
const SIGNATURE_FOLDER_NAME = 'ลายเซ็นเช็คชื่อ';

const TZ = 'Asia/Bangkok';
// =======================================================

const SHEET_ATT = 'เช็คชื่อ';
const SHEET_ROSTER = 'รายชื่อ';
const SHEET_SET = 'ตั้งค่า';

const ATT_HEAD = ['วันที่', 'เวลา', 'ชื่อ-นามสกุล', 'บริษัท/หน่วยงาน', 'สถานะ', 'สาย (นาที)', 'อุปกรณ์', 'ลายเซ็น', 'ไฟล์ลายเซ็น', 'id', 'timestamp', 'sigdata'];
const C = { date: 1, time: 2, name: 3, org: 4, status: 5, late: 6, device: 7, sig: 8, sigfile: 9, id: 10, ts: 11, data: 12 };
const ROSTER_HEAD = ['id', 'ชื่อ-นามสกุล', 'รหัส', 'ชื่อเล่น', 'บริษัท/หน่วยงาน'];
const SET_ROWS = [
  ['className', 'Claude AI', 'ชื่อคลาส'],
  ['startTime', '09:00', 'เวลาเริ่มเรียน (รูปแบบ HH:mm)'],
  ['lateMin', '15', 'ถือว่าสายเมื่อมาหลังเวลาเริ่มเกินกี่นาที'],
];
const DEFAULT_SETTINGS = { className: 'Claude AI', startTime: '09:00', lateMin: 15 };
const DEVICE_TH = { mobile: 'มือถือ', tablet: 'แท็บเล็ต', desktop: 'คอมพิวเตอร์' };

/** กด Run ฟังก์ชันนี้ 1 ครั้งหลังวางโค้ด เพื่อสร้างชีตและอนุญาตสิทธิ์ */
function setup() {
  ensure_();
  if (SAVE_SIGNATURE_TO_DRIVE) sigFolder_();
  bump_();
  Logger.log('พร้อมใช้งาน: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl());
  if (ADMIN_KEY === 'change-me-1234') Logger.log('⚠️ อย่าลืมเปลี่ยน ADMIN_KEY ก่อนใช้งานจริง');
}

// ----------------------- Web app -----------------------
function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const action = p.action || 'state';
    if (action === 'state') return json_(state_(p.date, p.v));
    if (action === 'ping') return json_(ping_(p.key));
    throw err_('bad_request', 'ไม่รู้จักคำสั่ง ' + action);
  } catch (err) {
    return json_({ ok: false, code: err.code || 'server', error: String(err.message || err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    let body;
    try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
    catch (x) { throw err_('bad_request', 'ข้อมูลที่ส่งมาไม่ใช่ JSON'); }
    try { lock.waitLock(25000); } catch (x) { throw err_('busy', 'ระบบกำลังยุ่ง ลองใหม่อีกครั้ง'); }
    switch (body.action) {
      case 'checkin': return json_(checkin_(body));
      case 'undo': checkKey_(body.key); return json_(undo_(body));
      case 'saveRoster': checkKey_(body.key); return json_(saveRoster_(body));
      case 'saveSettings': checkKey_(body.key); return json_(saveSettings_(body));
      default: throw err_('bad_request', 'ไม่รู้จักคำสั่ง ' + body.action);
    }
  } catch (err) {
    return json_({ ok: false, code: err.code || 'server', error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) { }
  }
}

/** เมื่อมีคนแก้ชีตด้วยมือ (เช่น เพิ่มรายชื่อ) ให้หน้าเว็บโหลดข้อมูลใหม่ */
function onEdit(e) {
  try { bump_(); } catch (x) { }
}

// ----------------------- Actions -----------------------
function state_(date, v) {
  const ver = ver_();
  if (v && String(v) === ver) return { ok: true, unchanged: true, version: ver };
  const sh = ensure_();
  const rows = readAtt_(sh.att);
  const sessions = unique_(rows.map(function (r) { return r.date; }).filter(Boolean));
  const checkins = {};
  if (date) {
    rows.forEach(function (r) {
      if (r.date === date && r.id && !checkins[r.id]) {
        checkins[r.id] = { sid: r.id, name: r.name, org: r.org, time: r.ts, sig: r.data, device: r.device };
      }
    });
  }
  return { ok: true, version: ver, roster: getRoster_(sh.ro), settings: getSettings_(sh.st), sessions: sessions, checkins: checkins };
}

function ping_(key) {
  const admin = !!key && String(key) === ADMIN_KEY;
  return { ok: true, admin: admin, version: ver_(), sheetUrl: admin ? SpreadsheetApp.getActiveSpreadsheet().getUrl() : '' };
}

function checkin_(b) {
  const date = String(b.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw err_('bad_request', 'วันที่ไม่ถูกต้อง');
  const rec = b.rec || {};
  const id = String(rec.sid || b.id || '');
  if (!id) throw err_('bad_request', 'ไม่มีรหัสผู้เรียน');

  const sh = ensure_();
  const person = getRoster_(sh.ro).filter(function (p) { return p.id === id; })[0];
  if (!person) throw err_('not_found', 'ไม่พบรายชื่อนี้ในชีต "รายชื่อ"');

  const existing = readAtt_(sh.att).filter(function (r) { return r.date === date && r.id === id; })[0];
  if (existing) {
    return { ok: true, duplicate: true, rec: { sid: id, name: existing.name, org: existing.org, time: existing.ts, sig: existing.data, device: existing.device } };
  }

  // ใช้เวลาจากเครื่องผู้เรียน (รองรับการเช็คชื่อตอนเน็ตหลุดแล้วส่งทีหลัง) แต่ถ้าคลาดเคลื่อนเกิน 12 ชม. ใช้เวลาเซิร์ฟเวอร์
  let ts = Number(rec.time);
  const now = Date.now();
  if (!ts || Math.abs(ts - now) > 12 * 3600 * 1000) ts = now;

  const set = getSettings_(sh.st);
  const start = Utilities.parseDate(date + ' ' + set.startTime, TZ, 'yyyy-MM-dd HH:mm').getTime();
  const diff = Math.round((ts - start) / 60000);
  const late = diff > set.lateMin;

  let sig = typeof rec.sig === 'string' ? rec.sig : '';
  if (sig.length > 200000 || !/^data:image\/png;base64,[A-Za-z0-9+\/=]+$/.test(sig)) sig = '';
  const device = DEVICE_TH[rec.device] || '';

  sh.att.appendRow([date, Utilities.formatDate(new Date(ts), TZ, 'HH:mm:ss'), person.name, person.org || '',
    late ? 'สาย' : 'ตรงเวลา', late ? diff : '', device, sig ? '✓ เซ็นแล้ว' : '-', '', id, ts, sig]);
  const row = sh.att.getLastRow();

  if (sig) {
    // แสดงรูปลายเซ็นในช่อง (ถ้า Google Sheet รองรับ) ไม่งั้นคงข้อความ "✓ เซ็นแล้ว"
    try {
      const img = SpreadsheetApp.newCellImage().setSourceUrl(sig).setAltTextTitle('ลายเซ็น ' + person.name).build();
      sh.att.getRange(row, C.sig).setValue(img);
      sh.att.setRowHeight(row, 48);
    } catch (x) { }
    if (SAVE_SIGNATURE_TO_DRIVE) {
      try {
        const blob = Utilities.newBlob(Utilities.base64Decode(sig.split(',')[1]), 'image/png', date + '_' + person.name + '.png');
        const file = sigFolder_().createFile(blob);
        sh.att.getRange(row, C.sigfile).setFormula('=HYPERLINK("' + file.getUrl() + '","เปิดไฟล์")');
      } catch (x) { }
    }
  }
  bump_();
  return { ok: true, rec: { sid: id, name: person.name, org: person.org || '', time: ts, sig: sig, device: device } };
}

function undo_(b) {
  const date = String(b.date || ''), id = String(b.id || '');
  const sh = ensure_();
  const rows = readAtt_(sh.att).filter(function (r) { return r.date === date && r.id === id; });
  rows.reverse().forEach(function (r) { sh.att.deleteRow(r.row); }); // ลบจากล่างขึ้นบน
  bump_();
  return { ok: true, removed: rows.length };
}

function saveRoster_(b) {
  const list = Array.isArray(b.students) ? b.students : null;
  if (!list) throw err_('bad_request', 'รายชื่อไม่ถูกต้อง');
  if (list.length > 5000) throw err_('bad_request', 'รายชื่อมากเกินไป');
  const sh = ensure_();
  const last = sh.ro.getLastRow();
  if (last > 1) sh.ro.getRange(2, 1, last - 1, ROSTER_HEAD.length).clearContent();
  const values = list
    .filter(function (s) { return s && String(s.name || '').trim(); })
    .map(function (s) {
      return [String(s.id || newId_()).slice(0, 80), str_(s.name, 200), str_(s.code, 80), str_(s.nick, 80), str_(s.org, 200)];
    });
  if (values.length) sh.ro.getRange(2, 1, values.length, ROSTER_HEAD.length).setValues(values);
  bump_();
  return { ok: true, count: values.length };
}

function saveSettings_(b) {
  const s = b.settings || {};
  const sh = ensure_();
  const vals = {
    className: str_(s.className, 80) || DEFAULT_SETTINGS.className,
    startTime: /^\d{1,2}:\d{2}$/.test(s.startTime || '') ? s.startTime : DEFAULT_SETTINGS.startTime,
    lateMin: String(Math.max(0, Math.min(240, parseInt(s.lateMin, 10) || 0))),
  };
  const data = sh.st.getRange(1, 1, Math.max(sh.st.getLastRow(), 1), 2).getValues();
  Object.keys(vals).forEach(function (k) {
    let row = -1;
    for (let i = 1; i < data.length; i++) if (String(data[i][0]) === k) { row = i + 1; break; }
    if (row < 0) { sh.st.appendRow([k, vals[k], '']); }
    else sh.st.getRange(row, 2).setValue(vals[k]);
  });
  bump_();
  return { ok: true };
}

// ----------------------- Sheets -----------------------
function ensure_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let att = ss.getSheetByName(SHEET_ATT);
  if (!att) att = ss.insertSheet(SHEET_ATT, 0);
  if (att.getLastRow() === 0) {
    att.getRange('A:B').setNumberFormat('@');
    att.getRange(1, 1, 1, ATT_HEAD.length).setValues([ATT_HEAD]).setFontWeight('bold').setBackground('#fde047');
    att.setFrozenRows(1);
    att.setColumnWidth(C.name, 220);
    att.setColumnWidth(C.org, 220);
    att.setColumnWidth(C.sig, 150);
    att.hideColumns(C.id, 3);
  }

  let ro = ss.getSheetByName(SHEET_ROSTER);
  if (!ro) ro = ss.insertSheet(SHEET_ROSTER);
  if (ro.getLastRow() === 0) {
    ro.getRange('A:E').setNumberFormat('@');
    ro.getRange(1, 1, 1, ROSTER_HEAD.length).setValues([ROSTER_HEAD]).setFontWeight('bold').setBackground('#fde047');
    ro.setFrozenRows(1);
    ro.setColumnWidth(2, 240);
    ro.setColumnWidth(5, 240);
    ro.getRange('A1').setNote('ปล่อยว่างได้ ระบบจะสร้าง id ให้อัตโนมัติ');
  }

  let st = ss.getSheetByName(SHEET_SET);
  if (!st) st = ss.insertSheet(SHEET_SET);
  if (st.getLastRow() === 0) {
    st.getRange('B:B').setNumberFormat('@');
    st.getRange(1, 1, 1, 3).setValues([['key', 'value', 'คำอธิบาย']]).setFontWeight('bold').setBackground('#fde047');
    st.getRange(2, 1, SET_ROWS.length, 3).setValues(SET_ROWS);
    st.setFrozenRows(1);
    st.setColumnWidth(3, 280);
  }
  return { att: att, ro: ro, st: st };
}

function readAtt_(att) {
  const last = att.getLastRow();
  if (last < 2) return [];
  const vals = att.getRange(2, 1, last - 1, ATT_HEAD.length).getValues();
  return vals.map(function (r, i) {
    return {
      row: i + 2,
      date: dateKey_(r[C.date - 1]),
      name: String(r[C.name - 1] || ''),
      org: String(r[C.org - 1] || ''),
      device: String(r[C.device - 1] || ''),
      id: String(r[C.id - 1] || ''),
      ts: Number(r[C.ts - 1]) || 0,
      data: String(r[C.data - 1] || ''),
    };
  });
}

function getRoster_(ro) {
  const last = ro.getLastRow();
  if (last < 2) return [];
  const range = ro.getRange(2, 1, last - 1, ROSTER_HEAD.length);
  const vals = range.getValues();
  let changed = false;
  const out = [];
  vals.forEach(function (r) {
    const name = String(r[1] || '').trim();
    if (!name) return;
    if (!String(r[0] || '').trim()) { r[0] = newId_(); changed = true; }
    out.push({ id: String(r[0]).trim(), name: name, code: String(r[2] || '').trim(), nick: String(r[3] || '').trim(), org: String(r[4] || '').trim() });
  });
  if (changed) range.setValues(vals); // เติม id ให้แถวที่เพิ่มเองในชีต
  return out;
}

function getSettings_(st) {
  const out = Object.assign({}, DEFAULT_SETTINGS);
  const last = st.getLastRow();
  if (last < 2) return out;
  st.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
    const k = String(r[0] || '').trim();
    let v = r[1];
    if (v instanceof Date) v = Utilities.formatDate(v, TZ, 'HH:mm');
    if (k === 'className' && String(v).trim()) out.className = String(v).trim();
    if (k === 'startTime' && /^\d{1,2}:\d{2}/.test(String(v))) out.startTime = String(v).slice(0, 5).replace(/^(\d):/, '0$1:');
    if (k === 'lateMin' && !isNaN(parseInt(v, 10))) out.lateMin = parseInt(v, 10);
  });
  return out;
}

function sigFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SIG_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (x) { } }
  const f = DriveApp.createFolder(SIGNATURE_FOLDER_NAME);
  props.setProperty('SIG_FOLDER', f.getId());
  return f;
}

// ----------------------- Helpers -----------------------
function checkKey_(key) {
  if (!key || String(key) !== ADMIN_KEY) throw err_('forbidden', 'รหัสผู้ดูแลไม่ถูกต้อง');
}
function ver_() { return PropertiesService.getScriptProperties().getProperty('VER') || '0'; }
function bump_() { PropertiesService.getScriptProperties().setProperty('VER', String(Date.now())); }
function newId_() { return 'r' + Utilities.getUuid().replace(/-/g, '').slice(0, 10); }
function str_(v, max) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max); }
function unique_(a) { const s = {}; return a.filter(function (x) { return s[x] ? false : (s[x] = true); }); }
function dateKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v || '').trim();
}
function err_(code, msg) { const e = new Error(msg); e.code = code; return e; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
