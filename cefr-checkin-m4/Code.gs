const SPREADSHEET_ID = '1UzYb-0nRTx0u9fogepiZ7DaTIc4LUjXaSuSgQ-FG-B0';
const STUDENT_SHEET = 'STUDENTS';
const CHECKIN_SHEET = 'CHECKIN';
const SETTINGS_SHEET = 'SETTINGS';
const TZ = 'Asia/Bangkok';

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = String(p.action || '').toLowerCase();
  if (action === 'lookup') return json_(lookup_(String(p.studentId || '').trim()));
  if (action === 'adminsummary') return json_(adminSummary_(String(p.pin || '')));
  if (action === 'adminroom') return json_(adminRoom_(String(p.pin || ''), String(p.room || ''), String(p.period || 'MORNING'), String(p.status || 'ALL')));
  return json_({ ok:true, service:'M.4 CEFR Prep Check-in 2569' });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(body.action || '').toLowerCase();
    if (action !== 'checkin') return json_({ok:false,message:'Invalid action'});
    return json_(checkin_(String(body.studentId || '').trim(), body));
  } catch (err) {
    return json_({ok:false,message:'ไม่สามารถประมวลผลคำขอได้'});
  }
}

function lookup_(studentId) {
  if (!studentId) return {ok:false,message:'กรุณากรอกรหัสประจำตัวนักเรียน'};
  const student = findStudent_(studentId);
  if (!student) return {ok:false,message:'ไม่พบรหัสนักเรียนนี้ในฐานข้อมูล ม.4'};

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
    message:prior ? 'ลงชื่อในรอบนี้แล้ว' : period.message
  };
}

function checkin_(studentId, payload) {
  if (!studentId) return {ok:false,message:'ไม่พบรหัสนักเรียน'};
  const student = findStudent_(studentId);
  if (!student) return {ok:false,message:'ไม่พบข้อมูลนักเรียน ม.4'};

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

function adminSummary_(pin) {
  const settings = readSettings_();
  if (pin !== String(settings.TEACHER_PIN || '1234')) return {ok:false,message:'รหัสสำหรับครูไม่ถูกต้อง'};
  const students = getAllStudents_();
  const logs = getAllCheckins_();
  const morning = new Set(logs.filter(x=>x.period==='MORNING').map(x=>x.studentId));
  const afternoon = new Set(logs.filter(x=>x.period==='AFTERNOON').map(x=>x.studentId));
  const roomMap = {};
  students.forEach(s=>{
    if (!roomMap[s.room]) roomMap[s.room] = {room:s.room,total:0,morning:0,afternoon:0};
    roomMap[s.room].total++;
    if (morning.has(s.studentId)) roomMap[s.room].morning++;
    if (afternoon.has(s.studentId)) roomMap[s.room].afternoon++;
  });
  return {
    ok:true,
    total:students.length,
    morning:morning.size,
    afternoon:afternoon.size,
    rooms:Object.keys(roomMap).sort((a,b)=>Number(a)-Number(b)).map(k=>roomMap[k]),
    updatedAt:Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss')
  };
}

function adminRoom_(pin, room, period, status) {
  const settings = readSettings_();
  if (pin !== String(settings.TEACHER_PIN || '1234')) return {ok:false,message:'รหัสสำหรับครูไม่ถูกต้อง'};
  period = period === 'AFTERNOON' ? 'AFTERNOON' : 'MORNING';
  status = ['CHECKED','UNCHECKED'].includes(status) ? status : 'ALL';

  const students = getAllStudents_().filter(s=>String(s.room)===String(room));
  const logs = getAllCheckins_().filter(x=>x.period===period);
  const logMap = {};
  logs.forEach(x=>{ if (!logMap[x.studentId]) logMap[x.studentId]=x; });

  let list = students.map(s=>({
    studentId:s.studentId,
    name:s.name,
    classRoom:s.classRoom,
    room:s.room,
    no:s.no,
    checked:!!logMap[s.studentId],
    checkedTime:logMap[s.studentId] ? logMap[s.studentId].time : ''
  }));
  if (status==='CHECKED') list=list.filter(x=>x.checked);
  if (status==='UNCHECKED') list=list.filter(x=>!x.checked);
  list.sort((a,b)=>(Number(a.no)||999)-(Number(b.no)||999));

  return {
    ok:true,
    room:String(room),
    period,
    periodLabel:period==='MORNING'?'รอบเช้า':'รอบบ่าย',
    status,
    total:students.length,
    checked:students.filter(s=>!!logMap[s.studentId]).length,
    unchecked:students.filter(s=>!logMap[s.studentId]).length,
    students:list
  };
}

function getAllStudents_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(STUDENT_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getRange(2,1,sh.getLastRow()-1,8).getDisplayValues();
  return values.filter(r=>String(r[0]).trim() && String(r[7]).toUpperCase() !== 'FALSE').map(r=>{
    const name = [r[1],r[2],r[3]].filter(Boolean).join('');
    const classRoom = r[5] ? `${r[4]}/${r[5]}` : r[4];
    return {studentId:String(r[0]).trim(),name,classRoom,room:String(r[5]).trim(),no:String(r[6]).trim()};
  });
}

function getAllCheckins_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(CHECKIN_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,12).getDisplayValues().map(r=>({
    studentId:String(r[1]).trim(),period:String(r[5]).trim(),time:String(r[8]).trim(),classRoom:String(r[3]).trim()
  })).filter(x=>x.studentId && x.period);
}

function findStudent_(studentId) {
  return getAllStudents_().find(s=>s.studentId===studentId) || null;
}

function findCheckin_(studentId, periodKey) {
  const found = getAllCheckins_().find(x=>x.studentId===studentId && x.period===periodKey);
  return found ? {time:found.time || ''} : null;
}

function getCurrentPeriod_() {
  const now = new Date();
  const ymd = Utilities.formatDate(now,TZ,'yyyy-MM-dd');
  const hhmm = Utilities.formatDate(now,TZ,'HH:mm');
  const settings = readSettings_();
  const eventDate = String(settings.EVENT_DATE || '2026-08-15');
  if (ymd !== eventDate) return {key:'',label:'',window:'',allowed:false,message:'ระบบลงชื่อเปิดเฉพาะวันที่ 15 สิงหาคม 2569'};

  const mStart = normalizeTime_(settings.MORNING_START || '07:30');
  const mEnd = normalizeTime_(settings.MORNING_END || '09:30');
  const aStart = normalizeTime_(settings.AFTERNOON_START || '12:30');
  const aEnd = normalizeTime_(settings.AFTERNOON_END || '14:30');

  if (hhmm >= mStart && hhmm <= mEnd) return {key:'MORNING',label:'รอบเช้า',window:`${mStart}–${mEnd} น.`,allowed:true,message:''};
  if (hhmm >= aStart && hhmm <= aEnd) return {key:'AFTERNOON',label:'รอบบ่าย',window:`${aStart}–${aEnd} น.`,allowed:true,message:''};
  return {key:'',label:'',window:'',allowed:false,message:'ขณะนี้อยู่นอกช่วงเวลาลงชื่อรอบเช้าและรอบบ่าย'};
}

function normalizeTime_(v) {
  const s = String(v || '').trim();
  const m = s.match(/^(\d{1,2}):(\d{2})/);
  return m ? `${m[1].padStart(2,'0')}:${m[2]}` : s;
}

function readSettings_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SETTINGS_SHEET);
  const obj = {};
  if (!sh || sh.getLastRow() < 2) return obj;
  sh.getRange(2,1,sh.getLastRow()-1,2).getDisplayValues().forEach(r=>{
    if (r[0]) obj[String(r[0]).trim()] = String(r[1]).trim();
  });
  return obj;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
