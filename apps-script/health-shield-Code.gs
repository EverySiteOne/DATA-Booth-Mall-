// HEALTH SHIELD – รับผลจากหน้าเกมบน GitHub แล้วบันทึกลงชีต "Health_Shield_Responses"
var SHEET_ID = '1wMzebZM4tdOHEtyLWe_TzEZBUuK_P0SQAcmjK9RD5go';
var SHEET = 'Health_Shield_Responses';
var HEADERS = ['Timestamp','Session ID','Language','Emergency Fund','Living Expenses','Health Reserve','Protection Learning','Total Allocated','Remaining','Event','Chosen Approach','Read Health Insurance','Read Critical Illness','Activity Score','Level','Consent Status'];
var EVENTS = ['Unexpected Medical Expense','Income Interruption','Critical Illness Awareness'];
var APPR = [['Use reserve','Review budget','Learn coverage','Review financial plan'],['Check reserve','Review expenses','Learn coverage','Re-prioritize expenses'],['Reserve fit','Learn differences','Read exclusions','Review family finances']];
var LEVELS = ['SHIELD MASTER','SHIELD BUILDER','SHIELD STARTER'];
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
  return out({ok:true, service:'health-shield'});
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
  var keys = ['e','l','h','p'], vals = [], total = 0;
  for (var i = 0; i < 4; i++) {
    var v = Number(d.al && d.al[keys[i]]);
    if (!isFinite(v) || v < 0 || v > START || v % 1) throw new Error('invalid amount');
    vals.push(v); total += v;
  }
  if (total > START || total < 1000) throw new Error('invalid total');
  var ev = Number(d.ev), pick = Number(d.pick);
  if (!(ev >= 0 && ev <= 2) || !(pick >= 0 && pick <= 3) || ev % 1 || pick % 1) throw new Error('invalid answer');
  var rh = d.rh === true, rc = d.rc === true;
  var score = 40 + 30 + ((rh || rc) ? 30 : 0);   // คำนวณซ้ำฝั่ง server
  var level = score >= 80 ? 0 : score >= 50 ? 1 : 2;
  var row = [new Date(), sid, d.lang === 'en' ? 'en' : 'th'].concat(vals, [total, START - total, EVENTS[ev], APPR[ev][pick], rh ? 'Yes' : 'No', rc ? 'Yes' : 'No', score, LEVELS[level], 'Agreed']);
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
