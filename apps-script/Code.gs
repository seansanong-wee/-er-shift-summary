/**
 * ระบบสรุปเวร ER — API บน Google Apps Script
 * ผูกกับชีต "สรุปเวรปีงบ 2569 (การตอบกลับ)" ใช้อ่าน เพิ่ม แก้ไข และลบแถวคำตอบฟอร์มจากหน้าเว็บ
 *
 * ตั้งค่า: Project Settings → Script Properties → เพิ่ม ACCESS_KEY = รหัสผ่านของทีม
 * พัฒนาโดย นายทวี แสนสนอง หัวหน้างานอุบัติเหตุฉุกเฉินและนิติเวช
 */

const SHEET_NAME = 'การตอบแบบฟอร์ม 1';
const TZ = 'Asia/Bangkok';

const NUM_KEYS = ['total', 'rn', 'revisit', 'resus', 'emerg', 'urgent', 'less', 'non', 'dw', 'admit', 'refer', 'dc',
  'refuse', 'dead', 'ett', 'hfnc', 'stroke', 'stemi', 'sepsis', 'shi', 'fastKid', 'cprOk', 'cprDead', 'legal',
  'accGen', 'accTraffic', 'autopsy', 'ems', 'pui', 'psych', 'ct'];
const TXT_KEYS = ['head', 'nurses', 'doctor', 'ep', 'recorder', 'maid', 'problem'];
const SHIFT_TEXT = { n: 'ดึก', m: 'เช้า', a: 'บ่าย' };

// header text → field key; order matters (first match wins)
const HRULES = [
  ['ts', /ประทับเวลา|timestamp/], ['email', /อีเมล|email/], ['date', /วัน\/เดือน|วันที่|^ว\/ด\/ป/],
  ['total', /ยอดรวม|ยอดผู้ป่วย/], ['revisit', /revisit/], ['rn', /ตรวจโดย/], ['resus', /resu/],
  ['less', /less/], ['non', /^non|nonurgent/], ['emerg', /emergen/], ['urgent', /urgent/],
  ['dw', /d\/w|ล้างแผล/], ['admit', /admit/], ['refer', /refer/], ['dc', /d\/c|จำหน่าย|ทุเลา/], ['refuse', /ไม่สมัครใจ/],
  ['cprOk', /cpr.*สำเร็จ/], ['cprDead', /cpr.*(เสีย|ตาย)/], ['dead', /dead|เสียชีวิต/],
  ['ett', /ett/], ['hfnc', /hfnc/], ['stroke', /stroke/], ['stemi', /stemi/], ['sepsis', /sepsis/],
  ['shi', /severehi|headinjury/], ['fastKid', /fast/], ['legal', /คดี/], ['accGen', /อุบัติเหตุทั่วไป/],
  ['accTraffic', /จราจร/], ['autopsy', /ชันสูตร/], ['ems', /ออกเหตุ|ems/], ['pui', /pui/], ['psych', /จิตเวช/],
  ['ct', /^ct/], ['head', /หน้าเวร|incharge/], ['ep', /แพทย์ep|^ep|staffep/], ['doctor', /แพทย์/],
  ['nurses', /พยาบาล|member/], ['recorder', /บันทึกข้อ|^data$/], ['maid', /แม่บ้าน|คนงาน/], ['problem', /ปัญหา/],
  ['shift', /^เวร/],
];

function doGet() {
  return out_({ ok: true, service: 'ระบบสรุปเวร ER', hint: 'ใช้ POST จากหน้าเว็บ' });
}

function doPost(e) {
  let p = {};
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (_) {}
  try {
    const key = PropertiesService.getScriptProperties().getProperty('ACCESS_KEY');
    if (!key) throw fail_('setup', 'ยังไม่ได้ตั้งรหัสผ่าน ACCESS_KEY ใน Script Properties');
    if (String(p.key || '') !== key) throw fail_('auth', 'รหัสผ่านไม่ถูกต้อง');
    switch (p.action) {
      case 'ping': return out_({ ok: true, title: SpreadsheetApp.getActiveSpreadsheet().getName() });
      case 'list': return out_(list_());
      case 'save': return out_(withLock_(() => save_(p)));
      case 'delete': return out_(withLock_(() => remove_(p)));
      default: throw fail_('bad_request', 'ไม่รู้จักคำสั่ง ' + p.action);
    }
  } catch (x) {
    return out_({ ok: false, code: x.code || 'error', message: x.message || String(x) });
  }
}

/* ---------- actions ---------- */

function list_() {
  const { sh, hdr, map } = layout_();
  const values = sh.getDataRange().getValues();
  const records = [];
  for (let r = 0; r < values.length; r++) {
    if (r === hdr) continue;
    const rec = toRecord_(values[r], map);
    if (rec) { rec.row = r + 1; records.push(rec); }
  }
  return { ok: true, title: SpreadsheetApp.getActiveSpreadsheet().getName(), records, at: new Date().toISOString() };
}

function save_(p) {
  const rec = p.rec || {};
  if (!/^\d{4}-\d{2}-\d{2}$/.test(rec.date || '')) throw fail_('bad_request', 'วันที่ไม่ถูกต้อง');
  if (!SHIFT_TEXT[rec.s]) throw fail_('bad_request', 'เวรไม่ถูกต้อง');
  const { sh, map } = layout_();
  const width = sh.getLastColumn();
  let row = Number(p.row) || 0;
  let cells;
  if (row) {
    cells = sh.getRange(row, 1, 1, width).getValues()[0];
    assertSame_(cells, map, p.expect);
  } else {
    cells = new Array(width).fill('');
    if (map.ts != null) cells[map.ts] = new Date();
  }
  const [y, m, d] = rec.date.split('-').map(Number);
  if (map.date != null) cells[map.date] = new Date(y, m - 1, d);
  if (map.shift != null) cells[map.shift] = SHIFT_TEXT[rec.s];
  NUM_KEYS.forEach(k => { if (map[k] != null && rec[k] != null) cells[map[k]] = Math.max(0, parseInt(rec[k], 10) || 0); });
  TXT_KEYS.forEach(k => { if (map[k] != null && rec[k] != null) cells[map[k]] = String(rec[k]); });
  if (row) sh.getRange(row, 1, 1, width).setValues([cells]);
  else { sh.appendRow(cells); row = sh.getLastRow(); }
  return { ok: true, row };
}

function remove_(p) {
  const rows = (p.rows || []).map(Number).filter(n => n > 1).sort((a, b) => b - a);
  if (!rows.length) throw fail_('bad_request', 'ไม่ได้ระบุแถวที่จะลบ');
  const { sh, map } = layout_();
  const width = sh.getLastColumn();
  // check every row first so a stale page never deletes the wrong shift
  rows.forEach(r => assertSame_(sh.getRange(r, 1, 1, width).getValues()[0], map, { date: p.date, s: p.s }));
  rows.forEach(r => sh.deleteRow(r));
  return { ok: true, deleted: rows.length };
}

/* ---------- helpers ---------- */

function layout_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const values = sh.getDataRange().getDisplayValues();
  for (let r = 0; r < values.length; r++) {
    const keys = values[r].map(headerKey_);
    if (keys.indexOf('date') < 0 || keys.indexOf('total') < 0) continue;
    const map = {};
    keys.forEach((k, c) => { if (k && map[k] == null) map[k] = c; });
    if (map.shift == null) map.shift = map.date + 1;
    return { sh, hdr: r, map };
  }
  throw fail_('layout', 'ไม่พบแถวหัวตารางที่มีคอลัมน์ วันที่ และ ยอดรวม ในชีต ' + sh.getName());
}

function headerKey_(s) {
  const h = String(s || '').toLowerCase().replace(/[\s.]/g, '');
  if (!h) return '';
  for (let i = 0; i < HRULES.length; i++) if (HRULES[i][1].test(h)) return HRULES[i][0];
  return '';
}

function toRecord_(row, map) {
  const ts = asDate_(row[map.ts]);
  const date = dayOf_(row[map.date], ts);
  const sv = String(row[map.shift] || '');
  const s = sv.indexOf('ดึก') >= 0 ? 'n' : sv.indexOf('เช้า') >= 0 ? 'm' : sv.indexOf('บ่าย') >= 0 ? 'a' : '';
  if (!date || !s) return null;
  const rec = { date, s, ts: ts ? Utilities.formatDate(ts, TZ, "yyyy-MM-dd'T'HH:mm:ssXXX") : '' };
  NUM_KEYS.forEach(k => { if (map[k] != null) rec[k] = parseInt(row[map[k]], 10) || 0; });
  TXT_KEYS.forEach(k => { rec[k] = map[k] != null ? clean_(row[map[k]]) : ''; });
  return rec;
}

function asDate_(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  const m = String(v || '').match(/^\s*(\d{1,2})\/(\d{1,2})\/(\d{4}),?\s+(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (!m) return null;
  let y = +m[3]; if (y > 2400) y -= 543;
  return new Date(y, m[2] - 1, +m[1], +m[4], +m[5], +(m[6] || 0));
}

function dayOf_(v, ts) {
  let y, mo, d;
  if (v instanceof Date && !isNaN(v)) {
    const p = Utilities.formatDate(v, TZ, 'yyyy-M-d').split('-').map(Number);
    y = p[0]; mo = p[1]; d = p[2];
  } else {
    const m = String(v || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (!m) return '';
    d = +m[1]; mo = +m[2]; y = +m[3];
  }
  if (y > 2400) y -= 543;
  if (y < 2000) {                 // years typed as 0064 / 64: use the year the answer was sent
    if (!ts) return '';
    y = ts.getFullYear();
    if (new Date(y, mo - 1, d) > ts) y--;
  }
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return '';
  return y + '-' + ('0' + mo).slice(-2) + '-' + ('0' + d).slice(-2);
}

function assertSame_(cells, map, expect) {
  if (!expect) return;
  const rec = toRecord_(cells, map);
  const ok = rec && rec.date === expect.date && rec.s === expect.s && (!expect.ts || rec.ts === expect.ts);
  if (!ok) throw fail_('conflict', 'แถวนี้ถูกแก้ไขหรือลบไปแล้วจากที่อื่น โหลดข้อมูลใหม่แล้วลองอีกครั้ง');
}

function clean_(v) {
  const t = String(v == null ? '' : v).trim();
  return ['-', '_', '--', '.', 'ไม่มี'].indexOf(t) >= 0 ? '' : t;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw fail_('busy', 'มีผู้ใช้อื่นกำลังบันทึก ลองใหม่อีกครั้ง');
  try { return fn(); } finally { lock.releaseLock(); }
}

function fail_(code, message) { const e = new Error(message); e.code = code; return e; }

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
