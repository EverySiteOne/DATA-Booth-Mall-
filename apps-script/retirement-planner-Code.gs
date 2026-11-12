// Google Apps Script – รับผลจากหน้าเกมบน GitHub แล้วบันทึกลง Sheet "Retirement_Responses"
var SHEET_ID = '1VGIibf5P7NpKKyFMcRoL4EuJQRRlZkdmpXVAZuaxtGU';
var SHEET = 'Retirement_Responses';
var HEADERS = ['Timestamp','Session ID','Current Age','Retirement Age','Life Expectancy','Desired Monthly Income','Current Retirement Savings','Years Until Retirement','Retirement Target','Monthly Saving Low Risk','Monthly Saving Medium Risk','Monthly Saving High Risk','Consent Status'];

function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function num(v,min,max){ v = Number(v); if (!isFinite(v) || v < min || v > max) throw new Error('invalid data'); return v; }

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
  return out({ok:true, service:'retirement-planner'});
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
  var row = [new Date(), sid, num(d.age,1,100), num(d.retAge,1,100), num(d.life,1,120),
    num(d.income,0,1e8), num(d.savings,0,1e11), num(d.years,1,100), num(d.target,0,1e13),
    num(d.low,0,1e11), num(d.mid,0,1e11), num(d.high,0,1e11), 'Agreed'];
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
