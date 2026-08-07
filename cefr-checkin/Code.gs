const SPREADSHEET_ID = '1zFeEJTJc16VV748Q6xc1vtm9_zdKCQm-g9qx8AphNXY';
const STUDENT_SHEET = 'STUDENTS';
const CHECKIN_SHEET = 'CHECKIN';
const SETTINGS_SHEET = 'SETTINGS';
const TZ = 'Asia/Bangkok';

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || '').toLowerCase();
  if (action === 'lookup') return json_(lookup_(String(e.parameter.studentId || '').trim()));
  return json_({ ok: true, service: 'M.5 CEFR Prep Check-in 2569' });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (String(body.action || '').toLowerCase() !== 'checkin') return json_({ok:false,message:'Invalid action'});
    return json_(checkin_(String(body.studentId || '').trim(), body));
  } catch (err) {
    return json_({ ok:false, message:'ไม่สามารถประมวลผลคำขอได้' });
  }
}

function lookup_(studentId) {
  if (!studentId) return {ok:false,message:'กรุณากรอกรหัสประจำตัวนักเรียน'};
  const student = findStudent_(studentId);
  if (!student) return {ok:false,message:'ไม่พบรหัสนักเรียนนี้ในฐานข้อมูล'};

  const period = getCurrentPeriod_();
  const prior = period.key ? findCheckin_(studentId, period.key) : null;

  return {
    ok:true,
    student,
    period:period.key,
    periodLabel:period.label,
    periodWindow:period.window,
    allowed:period.allowed && !prior,
    alreadyChecked:!!prior,
    checkedTime:prior ? prior.time : '',
    message: prior ? 'ลงชื่อในรอบนี้แล้ว' : period.message
  };
}

function checkin_(studentId, payload) {
  if (!studentId) return {ok:false,message:'ไม่พบรหัสนักเรียน'};
  const student = findStudent_(studentId);
  if (!student) return {ok:false,message:'ไม่พบข้อมูลนักเรียน'};

  const period = getCurrentPeriod_();
  if (!period.allowed || !period.key) return {ok:false,message:period.message || 'อยู่นอกช่วงเวลาลงชื่อ'};

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const prior = findCheckin_(studentId, period.key);
    if (prior) return {ok:false,message:`ลงชื่อ${period.label}แล้ว เวลา ${prior.time}`};

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(CHECKIN_SHEET);
    const now = new Date();
    const date = Utilities.formatDate(now, TZ, 'dd/MM/yyyy');
    const time = Utilities.formatDate(now, TZ, 'HH:mm:ss');
    sh.appendRow([
      now, student.studentId, student.name, student.classRoom, student.no,
      period.key, period.label, date, time, 'มา',
      String(payload && payload.source || ''), String(payload && payload.userAgent || '')
    ]);
    return {ok:true,period:period.key,periodLabel:period.label,checkinTime:time};
  } finally {
    lock.releaseLock();
  }
}

function findStudent_(studentId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(STUDENT_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  const values = sh.getRange(2,1,sh.getLastRow()-1,8).getDisplayValues();
  for (const r of values) {
    if (String(r[0]).trim() === studentId && String(r[7]).toUpperCase() !== 'FALSE') {
      const name = [r[1],r[2],r[3]].filter(Boolean).join('');
      const classRoom = r[5] ? `${r[4]}/${r[5]}` : r[4];
      return {studentId:r[0],name,classRoom,no:r[6]};
    }
  }
  return null;
}

function findCheckin_(studentId, periodKey) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(CHECKIN_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  const values = sh.getRange(2,1,sh.getLastRow()-1,12).getDisplayValues();
  for (const r of values) {
    if (String(r[1]).trim() === studentId && String(r[5]).trim() === periodKey) return {time:r[8] || ''};
  }
  return null;
}

function getCurrentPeriod_() {
  const now = new Date();
  const ymd = Utilities.formatDate(now,TZ,'yyyy-MM-dd');
  const hhmm = Utilities.formatDate(now,TZ,'HH:mm');
  const settings = readSettings_();
  const eventDate = settings.EVENT_DATE || '2026-08-08';
  if (ymd !== eventDate) return {key:'',label:'',window:'',allowed:false,message:'ระบบลงชื่อเปิดเฉพาะวันที่ 8 สิงหาคม 2569'};

  const mStart = settings.MORNING_START || '07:30';
  const mEnd = settings.MORNING_END || '09:00';
  const aStart = settings.AFTERNOON_START || '12:30';
  const aEnd = settings.AFTERNOON_END || '14:00';
  if (hhmm >= mStart && hhmm <= mEnd) return {key:'MORNING',label:'รอบเช้า',window:`${mStart}–${mEnd} น.`,allowed:true,message:''};
  if (hhmm >= aStart && hhmm <= aEnd) return {key:'AFTERNOON',label:'รอบบ่าย',window:`${aStart}–${aEnd} น.`,allowed:true,message:''};
  return {key:'',label:'',window:'',allowed:false,message:'ขณะนี้อยู่นอกช่วงเวลาลงชื่อรอบเช้าและรอบบ่าย'};
}

function readSettings_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SETTINGS_SHEET);
  const obj = {};
  if (!sh || sh.getLastRow() < 2) return obj;
  sh.getRange(2,1,sh.getLastRow()-1,2).getDisplayValues().forEach(r=>{ if(r[0]) obj[String(r[0]).trim()] = String(r[1]).trim(); });
  return obj;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
