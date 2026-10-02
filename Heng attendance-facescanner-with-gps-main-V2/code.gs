// ============================================================
//  GOOGLE APPS SCRIPT — REST API Backend
//  วิธีใช้: Deploy > New deployment > Web App
//           Execute as: Me | Who has access: Anyone
// ============================================================

function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : '';
  let result;

  if (action === 'getConfig') {
    result = getConfig();
  } else if (action === 'getKnownFaces') {
    result = getKnownFaces();
  } else {
    result = { error: 'Unknown action: ' + action };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e && e.postData ? e.postData.contents : '');
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Invalid JSON body' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const action = data.action;
  let result;

  if (action === 'registerUser') {
    result = registerUser(
      data.name,
      data.license,
      data.department,
      data.faceDescriptor
    );
  } else if (action === 'logAttendance') {
    result = logAttendance(data.name, data.license, data.department, data.lat, data.lng);
  } else if (action === 'saveConfig') {
    result = saveConfig(data.lat, data.lng, data.radius);
  } else {
    result = { error: 'Unknown action: ' + action };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getNextRow(sheet) {
  const lastRow = sheet.getLastRow();
  return Math.max(lastRow + 1, 2);
}

// --- ส่วนจัดการใบหน้า (Users) ---
function registerUser(name, license, department, faceDescriptor) {
  if (!name || !license || !department || !Array.isArray(faceDescriptor)) {
    return { success: false, error: 'ข้อมูลลงทะเบียนไม่ครบถ้วน' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Users');
  if (!sheet) sheet = ss.insertSheet('Users');

  sheet.getRange(1, 1, 1, 5).setValues([[
    'Name', 'License', 'Department', 'Face Descriptor', 'Registered Date'
  ]]);

  const nextRow = getNextRow(sheet);
  sheet.getRange(nextRow, 1, 1, 5).setValues([[
    name,
    "'" + license,
    department,
    JSON.stringify(faceDescriptor),
    new Date()
  ]]);

  return { success: true, message: 'บันทึกข้อมูลแพทย์และใบหน้าเรียบร้อย' };
}

function getKnownFaces() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(1, 1, lastRow, 5).getValues();

  let users = [];
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    const jsonStr = data[i][3];
    if (name && jsonStr) {
      try {
        users.push({ label: name, descriptor: JSON.parse(jsonStr) });
      } catch (e) {}
    }
  }
  return users;
}

// --- ส่วนบันทึกเวลา (Attendance) ---
function logAttendance(name, license, department, lat, lng) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Attendance');
  if (!sheet) sheet = ss.insertSheet('Attendance');

  sheet.getRange(1, 1, 1, 8).setValues([[
    'Name', 'License', 'Department', 'Time', 'Date', 'Latitude', 'Longitude', 'Google Map Link'
  ]]);

  const now = new Date();
  const mapLink = (lat && lng) ? `https://www.google.com/maps?q=${lat},${lng}` : '';
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'd/M/yyyy');
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');

  const nextRow = getNextRow(sheet);
  sheet.getRange(nextRow, 1, 1, 8).setValues([[
    name,
    license ? "'" + license : '',
    department || '',
    timeStr,
    "'" + dateStr,
    lat || '-',
    lng || '-',
    mapLink
  ]]);
  return { success: true, message: 'บันทึกเวลาสำเร็จ' };
}

// --- ส่วนจัดการ Config (GPS) ---
function saveConfig(lat, lng, radius) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config');

  if (!sheet) {
    sheet = ss.insertSheet('Config');
    sheet.getRange('A1:B1').setValues([['Parameter', 'Value']]);
    sheet.getRange('A2').setValue('Target Latitude');
    sheet.getRange('A3').setValue('Target Longitude');
    sheet.getRange('A4').setValue('Allowed Radius (KM)');
    sheet.setColumnWidth(1, 150);
  }

  sheet.getRange('B2').setValue(lat);
  sheet.getRange('B3').setValue(lng);
  sheet.getRange('B4').setValue(radius);

  return { success: true, message: 'บันทึกการตั้งค่าลง Google Sheets เรียบร้อย' };
}

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');

  let config = { lat: 0, lng: 0, radius: 0.5 };

  if (sheet) {
    const latVal = sheet.getRange('B2').getValue();
    const lngVal = sheet.getRange('B3').getValue();
    const radiusVal = sheet.getRange('B4').getValue();

    if (latVal !== '') config.lat = parseFloat(latVal);
    if (lngVal !== '') config.lng = parseFloat(lngVal);
    if (radiusVal !== '') config.radius = parseFloat(radiusVal);
  }

  return config;
}
