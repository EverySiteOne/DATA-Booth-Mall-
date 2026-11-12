// LIFE WALLET – รับผลจากหน้าเกมบน GitHub แล้วบันทึกลงชีต "Life_Wallet_Responses"
var SHEET_ID = '1KULwP_09cHuwV7j3tAauN42N7i3cdHHDCnB0Oi84pQ4';
var SHEET = 'Life_Wallet_Responses';
var HEADERS = ['Timestamp','Session ID','Language','Plan Name','Family & Home','Health & Medical','Emergency Fund','Education & Children','Savings & Retirement','Life & Income Protection','Unallocated','Life Event','Response Choice','Completeness Score','Consent Status'];
var KEYS = ['f','h','e','d','s','p'];
var NAMES = ['ครอบครัวและบ้าน','สุขภาพและค่ารักษา','เงินสำรองฉุกเฉิน','การศึกษาและอนาคตลูก','เงินออมและวัยเกษียณ','คุ้มครองชีวิตและรายได้'];
var EVENTS = ['Medical bill','Income pause','Family care'];
var CHOICES = ['Use emergency fund','Adjust plan / savings','Consider protection'];
var START = 100000;

function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// หน้าเกมบน GitHub ส่งข้อมูลมาที่นี่ (fetch POST)
function doPost(e) {
  try {
    var d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    return out(saveResult(d));
  } catch (err) {
    return out({ok:false, error:'bad request'});
  }
}

// เปิดลิงก์ Apps Script ตรง ๆ จะเห็นแค่ข้อความสถานะ (ตัวเกมอยู่บน GitHub)
function doGet() {
  return out({ok:true, service:'life-wallet'});
}

function saveResult(d) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    return save_(d);
  } catch (err) {
    return {ok:false, error:String(err.message || err)};
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function save_(d) {
  if (!d || d.consent !== 'Agreed') throw new Error('no consent');
  var sid = String(d.sessionId || '');
  if (!/^[A-Za-z0-9-]{8,64}$/.test(sid)) throw new Error('invalid session');
  var lang = d.lang === 'en' ? 'en' : 'th';
  var evi = Number(d.evi), pick = Number(d.pick);
  if (!(evi >= 0 && evi <= 2)) throw new Error('invalid event');
  if (!(pick >= -1 && pick <= 2)) throw new Error('invalid choice');

  var vals = [], total = 0, filled = 0, top = 0;
  for (var i = 0; i < KEYS.length; i++) {
    var v = Number(d.al && d.al[KEYS[i]]);
    if (!isFinite(v) || v < 0 || v > START || v % 1) throw new Error('invalid amount');
    vals.push(v); total += v;
    if (v > 0) filled++;
    if (v > vals[top]) top = i;
  }
  if (total > START || total <= 0) throw new Error('invalid total');
  var plan = 'แผนชีวิตเน้น' + NAMES[top];
  var score = Math.round(filled / 6 * 100);

  var row = [new Date(), sid, lang, plan].concat(vals, [START - total, EVENTS[evi], pick < 0 ? '-' : CHOICES[pick], score, 'Agreed']);

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); sh.setFrozenRows(1); }

  // กันข้อมูลซ้ำ ด้วย Session ID
  var last = sh.getLastRow();
  if (last > 1) {
    var ids = sh.getRange(2, 2, last - 1, 1).getValues();
    for (var j = 0; j < ids.length; j++) if (ids[j][0] === sid) return {ok:true, duplicate:true};
  }
  sh.appendRow(row);
  return {ok:true};
}