// HEART CATCHER – รับผลจากหน้าเกมบน GitHub แล้วบันทึกลงชีต "Heart_Catcher_Responses"
var SHEET_ID = '1UVZ5N7G3UfVdUiWhcFFBYPdSVSn0EwGwxPFvHKOTd-Y';
var SHEET = 'Heart_Catcher_Responses';
var HEADERS = ['Timestamp','Session ID','Language','Total Score','Items Grabbed','Grabs Used','Green Heart','Pink Heart','Fruit','Water Bottle','Health Shield','Bonus Coin','Mascot Toy','Title','Consent Status'];
var PTS = {hg:10, hp:15, fr:10, wt:10, sh:20, co:25, ms:30};
var COLS = ['hg','hp','fr','wt','sh','co','ms'];
var GRABS = 5, TIERS = [30, 70];
var TITLES = ['HEART BEGINNER','HEART COLLECTOR','HEART MASTER'];

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
  return out({ok:true, service:'heart-catcher'});
}

function saveResult(d) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); return save_(d); }
  catch (err) { return {ok:false, error:String(err.message || err)}; }
  finally { try { lock.releaseLock(); } catch (x) {} }
}

function save_(d) {
  if (!d || d.consent !== 'Agreed') throw new Error('no consent');
  var sid = String(d.sessionId || '');
  if (!/^[A-Za-z0-9-]{8,64}$/.test(sid)) throw new Error('invalid session');
  var c = d.c || {}, score = 0, n = 0, cnt = {};
  for (var i = 0; i < COLS.length; i++) {
    var k = COLS[i], v = Number(c[k]);
    if (!isFinite(v) || v < 0 || v > GRABS || v % 1) throw new Error('invalid count');
    cnt[k] = v; n += v; score += v * PTS[k];
  }
  var used = Number(d.used);
  if (!isFinite(used) || used < 0 || used > GRABS || used % 1) throw new Error('invalid used');
  if (n > used || Number(d.ok) !== n) throw new Error('inconsistent grabs');   // สำเร็จต้องไม่เกินจำนวนครั้งที่คีบ
  if (Number(d.score) !== score) throw new Error('score mismatch');           // คำนวณซ้ำฝั่ง server
  var tier = score >= TIERS[1] ? 2 : score >= TIERS[0] ? 1 : 0;
  var row = [new Date(), sid, d.lang === 'en' ? 'en' : 'th', score, n, used, cnt.hg, cnt.hp, cnt.fr, cnt.wt, cnt.sh, cnt.co, cnt.ms, TITLES[tier], 'Agreed'];
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); sh.setFrozenRows(1); }
  var last = sh.getLastRow();
  if (last > 1) {
    var ids = sh.getRange(2, 2, last - 1, 1).getValues();
    for (var j = 0; j < ids.length; j++) if (ids[j][0] === sid) return {ok:true, duplicate:true};
  }
  sh.appendRow(row);
  return {ok:true};
}
