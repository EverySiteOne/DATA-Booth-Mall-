/**
 * Booth Mall Insurance Dashboard 2569 — Google Apps Script backend
 * --------------------------------------------------------------
 * อ่านข้อมูลจริงทุกแท็บจาก Google Sheets แล้วส่งให้หน้าเว็บประมวลผล
 * - ไม่มี API Key / Service Account (สคริปต์อ่านชีตด้วยสิทธิ์ของผู้ Deploy)
 * - ไม่มีข้อมูลจำลอง: ถ้าอ่านไม่ได้จะส่ง error กลับไปแสดง
 */

const SPREADSHEET_ID = '1GfR4PGdqoxIwewf9dntILxaeaLNZokOqczz486coql0';
const APP_TITLE = 'Booth Mall Insurance 2569';
const MAX_ROWS = 5000;   // ป้องกันแท็บใหญ่ผิดปกติ (ถ้าเกินจะแจ้งว่า truncated)
const MAX_COLS = 80;

function doGet(e) {
  // ตรวจรหัสผ่านก่อนทุกครั้ง (ล็อกอิน / ตรวจ token / กันเปิดลิงก์ Apps Script ตรง ๆ)
  const guard = authGuard_(e); if (guard) return guard;

  // โหมด API: หน้าเว็บที่วางไว้ที่อื่น (เช่น GitHub Pages) ขอข้อมูลเป็น JSON — อ่านอย่างเดียว ไม่มีการแก้ไขชีต
  const api = e && e.parameter && e.parameter.api;
  if (api === 'getDashboardData' || api === 'getVersion') {
    const body = api === 'getDashboardData' ? getDashboardData() : getVersion();
    return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
  }
  // รองรับชื่อไฟล์ที่พิมพ์ต่างกันเล็กน้อย (Index / index / Index.html)
  const names = ['Index', 'index', 'Index.html', 'index.html'];
  let tpl = null;
  for (let i = 0; i < names.length && !tpl; i++) {
    try { tpl = HtmlService.createTemplateFromFile(names[i]); } catch (e) { /* ลองชื่อต่อไป */ }
  }
  if (!tpl) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:24px;line-height:1.6">' +
      '<h2>ไม่พบไฟล์ HTML ชื่อ Index</h2>' +
      '<p>ใน Apps Script กด <b>+ › HTML</b> แล้วพิมพ์ชื่อ <b>Index</b> (ไม่ต้องพิมพ์ .html) จากนั้นวางโค้ดจากไฟล์ Index.html แล้วกดบันทึก</p>' +
      '<p>แล้ว Deploy ใหม่: Deploy › Manage deployments › ✏️ › Version: New version › Deploy</p></div>'
    ).setTitle(APP_TITLE);
  }
  return tpl
    .evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSpreadsheet_() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    // กรณีผูกสคริปต์ไว้กับชีต (Extensions > Apps Script) ใช้ชีตที่เปิดอยู่แทน
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
    throw new Error('เปิดชีตไม่ได้ (ID: ' + SPREADSHEET_ID + ') — ตรวจสอบว่าบัญชีที่ Deploy มีสิทธิ์ดูชีตนี้: ' + e.message);
  }
}

/** เวลาแก้ไขชีตล่าสุด (ใช้เช็กว่าต้องโหลดใหม่หรือไม่) — เบามาก เรียกได้บ่อย */
function getVersion() {
  try {
    return JSON.stringify({ ok: true, version: DriveApp.getFileById(SPREADSHEET_ID).getLastUpdated().getTime() });
  } catch (e) {
    return JSON.stringify({ ok: false, version: null, error: String(e && e.message || e) });
  }
}

/** ข้อมูลดิบทุกแท็บ (ค่าจริง + ค่าที่แสดงในชีต) */
function getDashboardData() {
  try {
    const ss = getSpreadsheet_();
    const tz = ss.getSpreadsheetTimeZone() || 'Asia/Bangkok';
    const sheets = ss.getSheets().map(function (sh) {
      const lastRow = sh.getLastRow();
      const lastCol = sh.getLastColumn();
      const nr = Math.min(lastRow, MAX_ROWS);
      const nc = Math.min(lastCol, MAX_COLS);
      let values = [], display = [];
      if (nr > 0 && nc > 0) {
        const rg = sh.getRange(1, 1, nr, nc);
        values = rg.getValues().map(function (row) {
          return row.map(function (v) {
            return (v instanceof Date) ? { d: Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm:ss") } : v;
          });
        });
        display = rg.getDisplayValues();
      }
      return {
        name: sh.getName(),
        gid: sh.getSheetId(),
        hidden: sh.isSheetHidden(),
        lastRow: lastRow,
        lastCol: lastCol,
        truncated: lastRow > MAX_ROWS || lastCol > MAX_COLS,
        values: values,
        display: display
      };
    });

    let version = null;
    try { version = DriveApp.getFileById(ss.getId()).getLastUpdated().getTime(); } catch (e) { /* ไม่มีสิทธิ์ Drive: ใช้ refresh ตามเวลาแทน */ }

    return JSON.stringify({
      ok: true,
      spreadsheetName: ss.getName(),
      spreadsheetId: ss.getId(),
      url: ss.getUrl(),
      timeZone: tz,
      fetchedAt: new Date().toISOString(),
      version: version,
      sheets: sheets
    });
  } catch (e) {
    return JSON.stringify({ ok: false, error: String(e && e.message || e) });
  }
}

/** ใช้ทดสอบใน Editor: กด Run แล้วดู Execution log */
function testReadSheets() {
  const res = JSON.parse(getDashboardData());
  if (!res.ok) { Logger.log('ERROR: ' + res.error); return; }
  res.sheets.forEach(function (s) {
    Logger.log(s.name + ' (gid ' + s.gid + ') rows=' + s.lastRow + ' cols=' + s.lastCol +
      ' | header แถวแรก: ' + (s.display[0] || []).filter(String).join(' | '));
  });
}

/* =====================================================================
 * ระบบรหัสผ่าน — อ่านผู้ใช้จากแท็บ "ผู้ใช้งาน" อัตโนมัติ
 * คอลัมน์: รหัสผู้ใช้ | รหัสผ่าน | ชื่อ-นามสกุล | แผนก  (ถ้ามีคอลัมน์ "สถานะ" ใส่ ปิด/ระงับ = เข้าไม่ได้)
 * เพิ่ม/ลบ/แก้แถวในชีต = เพิ่ม/ลบ/เปลี่ยนรหัสทันที (ภายใน ~1 นาที)
 * ===================================================================== */
const AUTH = {
  SHEET_ID: '1GfR4PGdqoxIwewf9dntILxaeaLNZokOqczz486coql0',
  SHEET_NAME: 'ผู้ใช้งาน',
  SESSION_HOURS: 12,        // ล็อกอินค้างไว้กี่ชั่วโมง
  MAX_FAILS: 5,             // ใส่รหัสผิดได้กี่ครั้ง
  LOCK_MINUTES: 10,         // ล็อกบัญชีกี่นาทีเมื่อผิดเกินกำหนด
  SITE_URL: 'https://everysiteone.github.io/DATA-Booth-Mall-/'
};

/** เรียกจากบรรทัดแรกของ doGet(e) — คืนค่า output ถ้าต้องตอบเอง หรือ null ให้ doGet ทำงานต่อ */
function authGuard_(e) {
  const p = (e && e.parameter) || {};
  if (p.api === 'login') return authJson_(authLogin_(p.user, p.pass));
  if (p.api) {
    const who = authVerify_(p.token);
    if (!who) return authJson_({ ok: false, authRequired: true, error: 'เซสชันหมดอายุ หรือบัญชีถูกปิด กรุณาเข้าสู่ระบบอีกครั้ง' });
    return null; // token ถูกต้อง → ให้ doGet เดิมส่งข้อมูลตามปกติ
  }
  // เปิดลิงก์ Apps Script ตรง ๆ (ไม่มี ?api=) → ไม่แสดงข้อมูล ให้ไปหน้าเว็บที่มีล็อกอินแทน
  return HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;padding:40px;text-align:center">' +
    '<h3>Booth Mall 2569</h3><p>กรุณาเข้าใช้งานผ่านหน้าเว็บ</p>' +
    '<p><a href="' + AUTH.SITE_URL + '" target="_top">' + AUTH.SITE_URL + '</a></p></div>'
  ).setTitle('Booth Mall 2569');
}

/** รับการล็อกอินแบบ POST (รหัสผ่านไม่ไปอยู่ในลิงก์) — ถ้ามี doPost อยู่แล้ว ให้รวมเข้าด้วยกัน */
function doPost(e) {
  return authGuard_(e) || authJson_({ ok: false, error: 'คำขอไม่ถูกต้อง' });
}

function authLogin_(user, pass) {
  user = String(user || '').trim();
  pass = String(pass || '');
  if (!user || !pass) return { ok: false, error: 'กรุณากรอกรหัสผู้ใช้และรหัสผ่าน' };

  const cache = CacheService.getScriptCache();
  const failKey = 'authfail_' + user.toLowerCase();
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= AUTH.MAX_FAILS) return { ok: false, error: 'ใส่รหัสผิดหลายครั้ง กรุณารอ ' + AUTH.LOCK_MINUTES + ' นาทีแล้วลองใหม่' };

  const u = authFindUser_(user);
  if (!u || u.pass !== pass.trim()) {
    cache.put(failKey, String(fails + 1), AUTH.LOCK_MINUTES * 60);
    return { ok: false, error: 'รหัสผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  }
  cache.remove(failKey);
  const exp = Date.now() + AUTH.SESSION_HOURS * 3600 * 1000;
  return { ok: true, token: authSign_(u.user + '|' + exp), exp: exp, user: u.user, name: u.name, dept: u.dept };
}

/** ตรวจ token และเช็กว่าผู้ใช้ยังอยู่ในชีต (ลบแถว/เปลี่ยนรหัส/ปิดสถานะ = ตัดสิทธิ์ภายใน ~1 นาที) */
function authVerify_(token) {
  if (!token) return null;
  const i = token.lastIndexOf('.');
  if (i < 0) return null;
  const payload = authB64dec_(token.slice(0, i));
  if (authSign_(payload) !== token) return null;
  const parts = payload.split('|'); // user | exp | passFingerprint
  if (Number(parts[1]) < Date.now()) return null;
  const u = authFindUser_(parts[0]);
  if (!u || authFp_(u.pass) !== parts[2]) return null;
  return u;
}

function authFindUser_(user) {
  const key = String(user).trim().toLowerCase();
  return authUsers_().find(function (u) { return u.user.toLowerCase() === key; }) || null;
}

/** อ่านแท็บผู้ใช้งาน (cache 60 วินาที เพื่อความเร็ว) */
function authUsers_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('auth_users_v1');
  if (hit) return JSON.parse(hit);
  const sh = SpreadsheetApp.openById(AUTH.SHEET_ID).getSheetByName(AUTH.SHEET_NAME);
  if (!sh) throw new Error('ไม่พบแท็บ "' + AUTH.SHEET_NAME + '"');
  const rows = sh.getDataRange().getDisplayValues(); // ใช้ค่าที่แสดง เพื่อให้รหัส 0123 ไม่หายเลข 0
  const head = (rows.shift() || []).map(function (h) { return String(h).replace(/\s/g, ''); });
  const col = function (names, def) { for (var k = 0; k < names.length; k++) { var j = head.indexOf(names[k]); if (j >= 0) return j; } return def; };
  const cU = col(['รหัสผู้ใช้', 'ชื่อผู้ใช้', 'Username', 'User'], 0);
  const cP = col(['รหัสผ่าน', 'Password'], 1);
  const cN = col(['ชื่อ-นามสกุล', 'ชื่อนามสกุล', 'ชื่อ', 'Name'], 2);
  const cD = col(['แผนก', 'ฝ่าย', 'Department'], 3);
  const cS = col(['สถานะ', 'Status'], -1);
  const users = rows
    .filter(function (r) {
      if (!String(r[cU] || '').trim() || !String(r[cP] || '').trim()) return false;
      return !(cS >= 0 && /^(ปิด|ระงับ|ไม่ใช้งาน|inactive|off|no)$/i.test(String(r[cS]).trim()));
    })
    .map(function (r) {
      return { user: String(r[cU]).trim(), pass: String(r[cP]).trim(), name: String(r[cN] || '').trim(), dept: String(r[cD] || '').trim() };
    });
  cache.put('auth_users_v1', JSON.stringify(users), 60);
  return users;
}

/* ---------- token helpers ---------- */
function authSign_(payload) {
  // ผูกลายนิ้วมือรหัสผ่านไว้ใน token → เปลี่ยนรหัสในชีตแล้ว token เก่าใช้ไม่ได้
  if (payload.split('|').length === 2) {
    const u = authFindUser_(payload.split('|')[0]);
    payload += '|' + authFp_(u ? u.pass : '');
  }
  const sig = Utilities.computeHmacSha256Signature(payload, authSecret_());
  return Utilities.base64EncodeWebSafe(payload).replace(/=+$/, '') + '.' + Utilities.base64EncodeWebSafe(sig).replace(/=+$/, '');
}
function authFp_(pass) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature('fp:' + pass, authSecret_())).slice(0, 10);
}
function authB64dec_(s) {
  while (s.length % 4) s += '=';
  try { return Utilities.newBlob(Utilities.base64DecodeWebSafe(s)).getDataAsString(); } catch (e) { return ''; }
}
function authSecret_() {
  const props = PropertiesService.getScriptProperties();
  let k = props.getProperty('AUTH_SECRET');
  if (!k) { k = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('AUTH_SECRET', k); }
  return k;
}
function authJson_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** (ไม่บังคับ) รันฟังก์ชันนี้ใน Editor ถ้าต้องการบังคับให้ทุกคนออกจากระบบทันที */
function authLogoutEveryone() {
  PropertiesService.getScriptProperties().deleteProperty('AUTH_SECRET');
}
