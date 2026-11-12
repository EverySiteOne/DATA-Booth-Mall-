// HEALTH HERO DASH – รับผลจากหน้าเกมบน GitHub แล้วบันทึกลงชีต "Health_Hero_Dash_Responses"
var SHEET_ID = '1UdAYydEts1WsxquW0BCzti6enr1-Nrqp5xitQQs8UVk';
var SHEET = 'Health_Hero_Dash_Responses';
var HEADERS = ['Timestamp','Session ID','Language','Total Score','Hearts Collected','Coins','Water','Fruit','Shields','Bumps','Bonus Question','Bonus Correct','Item Score','Title','Consent Status'];
var PTS = {coin:5, heart:10, water:15, fruit:15, shield:20};
var BONUS = 50, TIERS = [450, 900];
var TITLES = ['Heart Collector','Full-Power Hero','Ultimate Health Guardian'];
var QUESTIONS = ['Daily health habits','Water and rest benefits','Reserve fund for unexpected health costs'];

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
  return out({ok:true, service:'health-hero-dash'});
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
  var c = d.c || {}, item = 0, cnt = {};
  for (var k in PTS) {
    var v = Number(c[k]);
    if (!isFinite(v) || v < 0 || v > 300 || v % 1) throw new Error('invalid count');
    cnt[k] = v; item += v * PTS[k];
  }
  var hits = Number(d.hits), qid = Number(d.qid);
  if (!isFinite(hits) || hits < 0 || hits > 200 || hits % 1) throw new Error('invalid hits');
  if (!(qid >= 0 && qid <= 2) || qid % 1) throw new Error('invalid question');
  var qok = d.qok === true, bonus = qok ? BONUS : 0, score = item + bonus;
  if (Number(d.score) !== score) throw new Error('score mismatch');   // คำนวณซ้ำฝั่ง server
  var tier = score >= TIERS[1] ? 2 : score >= TIERS[0] ? 1 : 0;
  var row = [new Date(), sid, d.lang === 'en' ? 'en' : 'th', score, cnt.heart, cnt.coin, cnt.water, cnt.fruit, cnt.shield, hits, QUESTIONS[qid], qok ? 'Yes' : 'No', item, TITLES[tier], 'Agreed'];
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
