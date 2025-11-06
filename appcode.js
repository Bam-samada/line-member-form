const SHEET_ID = '1f6WHZz1y_7OblNPLGSw3cSe8UY97vqmD9UzBNQTxS5g';
const SHEET_NAME = 'FormResponses';

// เรียกด้วย GET (เช่น เปิดลิงก์ใน browser ตรงๆ)
// เอาไว้เช็คว่า Web App ทำงาน + ตั้ง permission ถูก
function doGet(e) {
  return ContentService
    .createTextOutput('OK: Web App is running. Use POST to submit data.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    const req = parseRequest(e);

    // เตรียมชีต + หัวตาราง
    const sheet = ensureSheet();

    // เช็คซ้ำตาม LINE UserID
    if (req.action === 'check') {
      const exists = !!req.lineUserId && findByLineUserId(sheet, req.lineUserId) !== -1;
      return jsonOutput({ status: exists ? 'exists' : 'not_found', exists });
    }

    // สมัครจริง: ต้องมี lineUserId
    if (!req.lineUserId) {
      return jsonOutput({ status: 'error', message: 'Missing lineUserId' });
    }

    // ถ้ามีอยู่แล้ว ไม่บันทึกซ้ำ
    const rowIndex = findByLineUserId(sheet, req.lineUserId);
    if (rowIndex !== -1) {
      return jsonOutput({ status: 'exists', exists: true, message: 'Duplicate lineUserId' });
    }

    // บันทึกแถวใหม่
    sheet.appendRow([
      new Date(),
      req.lineUserId || '',
      req.lineDisplayName || '',
      req.company || '',
      req.fullName || '',
      req.nickname || '',
      req.phone || '',
      req.email || '',
      req.address || '',
      req.userAgent || ''
    ]);

    return jsonOutput({ status: 'success', message: 'Saved to Google Sheet' });

  } catch (error) {
    console.error('doPost error:', error);
    return jsonOutput({ status: 'error', message: error.toString() });
  }
}

// แยก parse request ให้รองรับทั้ง JSON / text/plain(JSON) / x-www-form-urlencoded
function parseRequest(e) {
  if (!e) return {};
  const p = e.parameter || {};
  const contentType = (e.postData && e.postData.type) || '';
  const body = (e.postData && e.postData.contents) || '';

  let fromBody = {};
  try {
    if (contentType.indexOf('application/json') !== -1 || contentType.indexOf('text/plain') !== -1) {
      fromBody = body ? JSON.parse(body) : {};
    }
  } catch (_) {
    // ถ้า parse ไม่ได้ จะใช้ค่าจาก form parameters แทน
  }

  const get = (k) => (fromBody[k] != null ? fromBody[k] : p[k]) || '';
  return {
    action: (get('action') || '').toString().trim() || 'submit',
    lineUserId: (get('lineUserId') || '').toString().trim(),
    lineDisplayName: (get('lineDisplayName') || '').toString().trim(),
    company: (get('company') || '').toString().trim(),
    fullName: (get('fullName') || '').toString().trim(),
    nickname: (get('nickname') || '').toString().trim(),
    phone: (get('phone') || '').toString().trim(),
    email: (get('email') || '').toString().trim(),
    address: (get('address') || '').toString().trim(),
    userAgent: (get('userAgent') || '').toString().trim()
  };
}

function ensureSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'Timestamp',
      'LINE UserID',
      'LINE Name',
      'Company',
      'Full Name',
      'Nickname',
      'Phone',
      'Email',
      'Address',
      'User Agent'
    ]);
  }
  return sheet;
}

// หา row index ของ LINE UserID (ไม่รวม header) ถ้าไม่พบ return -1
function findByLineUserId(sheet, lineUserId) {
  if (!lineUserId) return -1;
  const values = sheet.getDataRange().getValues();
  // คอลัมน์ที่ 2 (index 1) คือ LINE UserID ตามหัวข้อที่ตั้งไว้
  for (let i = 1; i < values.length; i++) {
    if ((values[i][1] || '').toString().trim() === lineUserId) {
      return i + 1; // row index (1-based)
    }
  }
  return -1;
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
