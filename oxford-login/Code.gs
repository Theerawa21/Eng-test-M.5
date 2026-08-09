const SPREADSHEET_ID = '1FtjWFjdwxY38bKNruNuJxkYY5RJ2hGgxsisPtL-VqFQ';
const DATABASE_SHEET = 'DATABASE';
const EXAM_URL = 'https://www.oxfordenglishtesting.com';

function doGet(e) {
  const studentId = String((e && e.parameter && e.parameter.id) || '').trim();
  const callback = String((e && e.parameter && e.parameter.callback) || '').trim();

  const payload = lookupStudent_(studentId);
  const json = JSON.stringify(payload);

  if (callback && /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback)) {
    return ContentService
      .createTextOutput(`${callback}(${json});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function lookupStudent_(studentId) {
  if (!studentId) {
    return { found: false, message: 'กรุณากรอกรหัสประจำตัวนักเรียน' };
  }

  if (!/^\d{3,20}$/.test(studentId)) {
    return { found: false, message: 'รูปแบบรหัสประจำตัวนักเรียนไม่ถูกต้อง' };
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(DATABASE_SHEET);

  if (!sheet) {
    return { found: false, message: 'ไม่พบฐานข้อมูลนักเรียน' };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { found: false, message: 'ยังไม่มีข้อมูลนักเรียนในระบบ' };
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getDisplayValues();
  const row = data.find(r => String(r[0] || '').trim() === studentId);

  if (!row) {
    return { found: false, message: 'ไม่พบข้อมูล กรุณาตรวจสอบรหัสประจำตัวนักเรียนอีกครั้ง' };
  }

  return {
    found: true,
    studentId: row[0],
    name: row[1],
    room: row[2],
    studentNo: row[3],
    group: row[4],
    username: row[5],
    password: row[6],
    orgId: row[7] || '103713',
    examUrl: EXAM_URL
  };
}
