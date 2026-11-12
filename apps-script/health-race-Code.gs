// HEALTH RACE – รับผลจากหน้าเกมบน GitHub แล้วบันทึกลงชีต "Health_Race_Responses"
var SHEET_ID = '1MUYwcCM_LOyBVBu7ATAULYJJUG0UsileUbTDcWGyXuY';
var SHEET = 'Health_Race_Responses';
var HEADERS = ['Timestamp','Session ID','Language','Car','Total Score','Items Collected','Green Heart','Fruit','Water Bottle','Health Shield','Bonus Coin','Distance (m)','Bumps','Boosts','Title','Consent Status'];
var PTS = {he:10, fr:10, wt:10, sh:20, co:25};
var COLS = ['he','fr','wt','sh','co'];
var CARS = ['Lime','Sky','Orange'];
var TIERS = [350, 600];
var TITLES = ['HEART RIDER','SHIELD RACER','HEALTH RACE MASTER'];

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
  return out({ok:true, service:'health-race'});
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
    if (!isFinite(v) || v < 0 || v > 400 || v % 1) throw new Error('invalid count');
    cnt[k] = v; n += v; score += v * PTS[k];
  }
  var car = Number(d.car), dist = Number(d.dist), hits = Number(d.hits), boosts = Number(d.boosts);
  if (!(car >= 0 && car <= 2) || car % 1) throw new Error('invalid car');
  if (!isFinite(dist) || dist < 0 || dist > 5000 || dist % 1) throw new Error('invalid distance');
  if (!isFinite(hits) || hits < 0 || hits > 200 || hits % 1) throw new Error('invalid hits');
  if (!isFinite(boosts) || boosts < 0 || boosts > 30 || boosts % 1) throw new Error('invalid boosts');
  if (Number(d.score) !== score) throw new Error('score mismatch');   // คำนวณซ้ำฝั่ง server
  var tier = score >= TIERS[1] ? 2 : score >= TIERS[0] ? 1 : 0;
  var row = [new Date(), sid, d.lang === 'en' ? 'en' : 'th', CARS[car], score, n, cnt.he, cnt.fr, cnt.wt, cnt.sh, cnt.co, dist, hits, boosts, TITLES[tier], 'Agreed'];
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
