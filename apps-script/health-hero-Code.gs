// HEALTH HERO – รับผลจากหน้าเกมบน GitHub แล้วบันทึกลงชีต "Health_Hero_Responses"
var SHEET_ID = '1Bjexc9b3x2Y0oY2Dit_cmLs5dXSLPafZ9VU6XdLAzhU';
var SHEET = 'Health_Hero_Responses';
var HEADERS = ['Timestamp','Session ID','Language','Habit 1','Habit 2','Scenario','Challenge Choice','Read Health Insurance','Read Critical Illness','Activity Score','Level','Consent Status'];
var HABITS = {m:'Move More', e:'Eat Well', s:'Sleep & Rest', c:'Health Check'};
var SCEN = ['Unexpected Medical Expense','Income Interruption','Family Health Planning'];
var TYPES = [['Reserve','Review budget','Learn coverage','Not prepared yet'],['Reserve','Review expenses','Learn coverage','Not prepared yet'],['Health check','Reserve','Learn coverage','Review family plan']];
var LEVELS = ['HEALTH HERO','HERO IN PROGRESS','START YOUR JOURNEY'];

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
  return out({ok:true, service:'health-hero'});
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
  var h = d.habits;
  if (!h || h.length !== 2 || h[0] === h[1] || !HABITS[h[0]] || !HABITS[h[1]]) throw new Error('invalid habits');
  var sc = Number(d.sc), pick = Number(d.pick);
  if (!(sc >= 0 && sc <= 2) || !(pick >= 0 && pick <= 3) || sc % 1 || pick % 1) throw new Error('invalid answer');
  var rh = d.rh === true, rc = d.rc === true;
  var score = 40 + 30 + ((rh || rc) ? 30 : 0);   // คำนวณซ้ำฝั่ง server
  var level = score >= 80 ? 0 : score >= 50 ? 1 : 2;
  var row = [new Date(), sid, d.lang === 'en' ? 'en' : 'th', HABITS[h[0]], HABITS[h[1]], SCEN[sc], TYPES[sc][pick], rh ? 'Yes' : 'No', rc ? 'Yes' : 'No', score, LEVELS[level], 'Agreed'];
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); sh.setFrozenRows(1); }
  var last = sh.getLastRow();
  if (last > 1) {
    var ids = sh.getRange(2, 2, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) if (ids[i][0] === sid) return {ok:true, duplicate:true};
  }
  sh.appendRow(row);
  return {ok:true};
}
