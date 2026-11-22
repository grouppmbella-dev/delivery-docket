const SHEET_ID = '1w-vEijSeHI5rYtxjmtxyqv2Eh1P5yl8ShZHZlIftnOg';
const SHEET_NAME = 'DD';
const NUM_COLS = 30;

const TOKEN_DAYS = 30;   // login ครั้งเดียวใช้ได้กี่วัน

// ===================== AUTH =====================
// รหัสผ่านเก็บใน Script Properties (ไม่อยู่ในโค้ด/ไม่อยู่ใน GitHub)
// ตั้งค่า: Project Settings (⚙️) → Script Properties → เพิ่ม  DD_PASSWORD = รหัสใหม่
function getSecret() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('DD_SECRET');
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('DD_SECRET', secret);
  }
  return secret;
}

function sign(payload) {
  const sig = Utilities.computeHmacSha256Signature(payload, getSecret(), Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(sig);
}

function login(data) {
  const password = PropertiesService.getScriptProperties().getProperty('DD_PASSWORD');
  if (!password) return { success: false, error: 'ยังไม่ได้ตั้งรหัสผ่าน (DD_PASSWORD) ใน Script Properties' };
  const name = String(data.name || '').trim();
  if (!name) return { success: false, error: 'Please enter your name.' };
  if (String(data.password || '') !== password) {
    Utilities.sleep(1500);   // หน่วงเวลา กันการสุ่มรหัส
    return { success: false, error: 'Incorrect password. Please try again.' };
  }
  const exp = Date.now() + TOKEN_DAYS * 24 * 60 * 60 * 1000;
  const payload = Utilities.base64EncodeWebSafe(name + '|' + exp, Utilities.Charset.UTF_8);
  return { success: true, token: payload + '.' + sign(payload), name: name };
}

// คืนชื่อผู้ใช้ถ้า token ถูกต้องและยังไม่หมดอายุ ไม่งั้นคืน null
function verifyToken(token) {
  try {
    const parts = String(token || '').split('.');
    if (parts.length !== 2) return null;
    if (sign(parts[0]) !== parts[1]) return null;
    const text = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString('UTF-8');
    const i = text.lastIndexOf('|');
    const exp = Number(text.slice(i + 1));
    if (!exp || Date.now() > exp) return null;
    return text.slice(0, i);
  } catch (e) { return null; }
}

const AUTH_FAIL = { success: false, authError: true, error: 'Please sign in again.' };

// ===================== ROUTER =====================
function doPost(e) {
  let data;
  try { data = JSON.parse(e.postData.contents); }
  catch (err) { return response({ success: false, error: 'Bad request' }); }

  if (data.action === 'login') return response(login(data));
  if (!verifyToken(data.token)) return response(AUTH_FAIL);

  // ล็อกให้ทำทีละคำขอ กันกด Save ซ้ำ / หลายคนบันทึกพร้อมกันแล้วข้อมูลชนกัน
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return response({ success: false, error: 'ระบบกำลังบันทึกงานอื่นอยู่ กรุณาลองใหม่อีกครั้ง' });
  }
  try {
    const action = data.action;
    let result;
    if (action === 'createDocket') result = createDocket(data);
    else if (action === 'updateSigned') result = updateSigned(data);
    else if (action === 'updateDocket') result = updateDocket(data);
    else if (action === 'deleteItem') result = deleteItem(data);
    else if (action === 'uploadImage') result = uploadImage(data);
    else result = { success: false, error: 'Unknown action: ' + action };
    SpreadsheetApp.flush();
    return response(result);
  } catch (err) {
    return response({ success: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const action = e.parameter.action;
  if (!verifyToken(e.parameter.token)) return response(AUTH_FAIL);
  if (action === 'getDockets') return response(getDockets());
  if (action === 'getDocketById') return response(getDocketById(e.parameter.docNo));
  if (action === 'getLogoBase64') return response(getLogoBase64());
  if (action === 'fixPictureFileIds') return response(fixPictureFileIds());
  return response({ success: false, error: 'Unknown action: ' + action });
}

// ===================== HELPERS =====================
function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
}

// เทียบเลข Doc แบบตัดช่องว่าง กันกรณีมีเว้นวรรคติดมาแล้วหาไม่เจอ
function sameDoc(a, b) {
  return String(a || '').trim() === String(b || '').trim();
}

function readAll(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, NUM_COLS).getValues();
}

// อัปโหลดรูป (ถ้ามี) แล้วคืน File ID — ทำ "ก่อน" แตะชีต
function resolveFileId(item) {
  // 1) หน้าเว็บอัปโหลดสำเร็จแล้ว -> ใช้ fileId นั้น
  if (item.fileId) return item.fileId;
  // 2) มีรูปใหม่แต่หน้าเว็บอัปโหลดไม่สำเร็จ -> อัปโหลดที่นี่แทน
  if (item.imageData && item.fileName) {
    try {
      const r = uploadImage({
        imageData: item.imageData,
        mimeType: item.mimeType || 'image/jpeg',
        fileName: item.fileName
      });
      if (r.success) return r.fileId;
    } catch (e) { Logger.log('Image upload failed: ' + e.message); }
  }
  // 3) ไม่มีรูปใหม่ -> ใช้รูปเดิม (File ID) ตอน Edit
  return item.picture || '';
}

// สร้างข้อมูล 1 แถว (old = แถวเดิม ใช้เก็บลายเซ็น/สถานะไว้)
function buildRow(rowNo, docNo, data, item, index, fileId, old) {
  old = old || [];
  return [
    rowNo,
    docNo,
    data.requestDate,
    data.project,
    data.jobNumber,
    data.sendTo,
    index + 1,
    item.description,
    item.qty,
    item.unit,
    fileId,
    item.remark || '',
    data.nameIssueBy,
    data.issueDate,
    old[14] || '', old[15] || '',
    data.nameCheckBy,
    data.checkDate,
    old[18] || '', old[19] || '',
    data.nameDeliveryBy,
    data.deliveryDate,
    old[22] || '', old[23] || '',
    data.carNumberPlate,
    data.nameReceivedBy,
    data.receivedDate,
    old[27] || '', old[28] || '',
    old[29] || 'Draft'
  ];
}

function getNextDocNo() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  const year = new Date().getFullYear();
  if (lastRow <= 1) return 'DDPM/' + year + '/0001';
  const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const docNos = data.map(r => String(r[0]).trim()).filter(v => v !== '');
  const unique = [...new Set(docNos)];
  const thisYear = unique.filter(d => d.includes('/' + year + '/'));
  if (thisYear.length === 0) return 'DDPM/' + year + '/0001';
  const nums = thisYear.map(d => parseInt(d.split('/')[2], 10)).filter(n => !isNaN(n));
  const next = (nums.length ? Math.max(...nums) : 0) + 1;
  return 'DDPM/' + year + '/' + String(next).padStart(4, '0');
}

function getNextRowNo() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const count = data.filter(r => String(r[0]).trim().startsWith('DDPM/')).length;
  return count + 1;
}

// ===================== CREATE =====================
function createDocket(data) {
  const sheet = getSheet();
  const items = data.items || [];
  if (items.length === 0) return { success: false, error: 'No items' };

  // 1) อัปโหลดรูปให้เสร็จก่อน
  const fileIds = items.map(resolveFileId);

  // 2) ออกเลข + เขียนลงชีตครั้งเดียว (อยู่ใต้ lock จึงไม่ได้เลขซ้ำ)
  const docNo = getNextDocNo();
  let rowNo = getNextRowNo();
  const rows = items.map((item, i) => buildRow(rowNo++, docNo, data, item, i, fileIds[i], null));
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, NUM_COLS).setValues(rows);

  return { success: true, docNo: docNo };
}

// ===================== READ =====================
const HEADERS = ['docNoNo','docNo','requestDate','project','jobNumber','sendTo',
  'itemNo','description','qty','unit','picture','remark',
  'nameIssueBy','issueDate','issueSigned','issueSignedDate',
  'nameCheckBy','checkDate','checkSigned','checkSignedDate',
  'nameDeliveryBy','deliveryDate','deliverySigned','deliverySignedDate',
  'carNumberPlate','nameReceivedBy','receivedDate','receivedSigned',
  'receivedSignedDate','status'];

function rowToObj(row) {
  const obj = {};
  HEADERS.forEach((h, i) => obj[h] = row[i]);
  obj.docNo = String(obj.docNo).trim();
  return obj;
}

function getDockets() {
  const data = readAll(getSheet());
  const result = data
    .filter(row => String(row[1] || '').trim().startsWith('DDPM/'))
    .map(rowToObj);
  return { success: true, data: result };
}

function getDocketById(docNo) {
  const data = readAll(getSheet());
  const items = data
    .filter(row => sameDoc(row[1], docNo) && String(row[1]).trim().startsWith('DDPM/'))
    .map(rowToObj);
  return { success: true, data: items };
}

// ===================== UPDATE (แก้ปัญหาค้างแล้วหาย) =====================
// เดิม: ลบแถวเก่าทิ้งทั้งหมดก่อน -> อัปโหลดรูป -> ค่อยแทรกแถวใหม่ทีละแถว
//       ถ้าหลุด/timeout/กดซ้ำกลางทาง = แถวถูกลบไปแล้วแต่ยังไม่ได้เขียนใหม่ -> Docket หาย
// ใหม่: อัปโหลดรูปก่อน -> เขียนทับแถวเดิม -> แล้วค่อยเพิ่ม/ลบส่วนต่าง
//       ถ้าพังกลางทาง ข้อมูลเดิมยังอยู่เสมอ
function updateDocket(data) {
  const sheet = getSheet();
  const docNo = String(data.docNo || '').trim();
  const items = data.items || [];
  if (!docNo) return { success: false, error: 'Missing docNo' };
  if (items.length === 0) return { success: false, error: 'No items' };

  // 1) อัปโหลดรูปทั้งหมดก่อน (ยังไม่แตะชีต)
  const fileIds = items.map(resolveFileId);

  // 2) อ่านชีตหลังอัปโหลดเสร็จ (ได้ข้อมูลล่าสุด)
  const allData = readAll(sheet);
  const rowIndices = [];
  allData.forEach((row, i) => { if (sameDoc(row[1], docNo)) rowIndices.push(i); });

  // 3) ถ้าหาไม่เจอ (เช่นหายไปจากบั๊กเดิม) -> กู้คืนด้วยข้อมูลบนหน้าจอ ใช้เลขเดิม
  if (rowIndices.length === 0) {
    let rowNo = getNextRowNo();
    const rows = items.map((item, i) => buildRow(rowNo++, docNo, data, item, i, fileIds[i], null));
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, NUM_COLS).setValues(rows);
    return { success: true, restored: true };
  }

  const existingFirst = allData[rowIndices[0]];
  const newRows = items.map((item, i) => {
    const oldRow = allData[rowIndices[i]];             // แถวเดิมตำแหน่งเดียวกัน (ถ้ามี)
    const rowNo = oldRow ? oldRow[0] : existingFirst[0];
    return buildRow(rowNo, docNo, data, item, i, fileIds[i], existingFirst);
  });

  const oldCount = rowIndices.length;
  const newCount = newRows.length;
  const overlap = Math.min(oldCount, newCount);

  // 4) เขียนทับแถวเดิม
  const contiguous = rowIndices.every((v, k) => k === 0 || v === rowIndices[k - 1] + 1);
  if (contiguous) {
    sheet.getRange(rowIndices[0] + 2, 1, overlap, NUM_COLS).setValues(newRows.slice(0, overlap));
  } else {
    for (let k = 0; k < overlap; k++) {
      sheet.getRange(rowIndices[k] + 2, 1, 1, NUM_COLS).setValues([newRows[k]]);
    }
  }

  // 5a) รายการเพิ่มขึ้น -> แทรกต่อท้ายแถวสุดท้ายของ Docket
  if (newCount > oldCount) {
    const afterRow = rowIndices[oldCount - 1] + 2;
    const extra = newRows.slice(oldCount);
    sheet.insertRowsAfter(afterRow, extra.length);
    sheet.getRange(afterRow + 1, 1, extra.length, NUM_COLS).setValues(extra);
  }

  // 5b) รายการลดลง -> ลบแถวส่วนเกิน (ล่างขึ้นบน)
  if (newCount < oldCount) {
    for (let k = oldCount - 1; k >= newCount; k--) {
      sheet.deleteRow(rowIndices[k] + 2);
    }
  }

  return { success: true };
}

// ===================== DELETE ITEM =====================
function deleteItem(data) {
  const sheet = getSheet();
  const allData = readAll(sheet);

  const docketRows = allData.map((row, i) => ({ row, i })).filter(r => sameDoc(r.row[1], data.docNo));
  if (docketRows.length === 0) return { success: false, error: 'Docket not found' };
  if (docketRows[0].row[29] !== 'Draft') return { success: false, error: 'Only Draft dockets can be deleted' };

  const target = docketRows.find(r => String(r.row[6]) === String(data.itemNo));
  if (!target) return { success: false, error: 'Item not found' };

  sheet.deleteRow(target.i + 2);

  // เลขลำดับใหม่ของรายการที่เหลือ
  let itemNum = 1;
  let remaining = 0;
  docketRows.forEach(r => {
    if (r.i === target.i) return;
    const sheetRow = r.i > target.i ? r.i + 1 : r.i + 2;   // แถวที่อยู่ใต้แถวที่ลบ เลื่อนขึ้น 1
    sheet.getRange(sheetRow, 7).setValue(itemNum++);
    remaining++;
  });

  return { success: true, remainingItems: remaining };
}

// ===================== SIGN =====================
function updateSigned(data) {
  const sheet = getSheet();
  const allData = readAll(sheet);
  if (allData.length === 0) return { success: false, error: 'No data' };
  const colMap = { 'issue': [15, 16], 'check': [19, 20], 'delivery': [23, 24], 'received': [29, 30] };
  const statusMap = { 'issue': 'Issued', 'delivery': 'Delivered', 'received': 'Completed' };
  const cols = colMap[data.type];
  if (!cols) return { success: false, error: 'Unknown sign type' };
  const today = new Date().toLocaleDateString('en-GB');
  let found = 0;
  allData.forEach((row, index) => {
    if (sameDoc(row[1], data.docNo)) {
      sheet.getRange(index + 2, cols[0], 1, 2).setValues([['Signed', today]]);
      if (statusMap[data.type]) sheet.getRange(index + 2, 30).setValue(statusMap[data.type]);
      found++;
    }
  });
  if (found === 0) return { success: false, error: 'Docket not found' };
  return { success: true };
}

// ===================== IMAGES =====================
function uploadImage(data) {
  const folder = getDDImagesFolder();
  const blob = Utilities.newBlob(
    Utilities.base64Decode(data.imageData),
    data.mimeType,
    data.fileName
  );
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { success: true, fileId: file.getId(), fileName: data.fileName };
}

function getFileIdByName(fileName) {
  try {
    const folder = getDDImagesFolder();
    const files = folder.getFilesByName(fileName);
    if (files.hasNext()) {
      const file = files.next();
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return file.getId();
    }
    return null;
  } catch (e) { return null; }
}

function fixPictureFileIds() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) return { success: true, fixed: 0 };
  const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  let fixed = 0;
  data.forEach((row, i) => {
    const pic = String(row[10] || '').trim();
    if (pic && pic.includes('.') && pic.length < 50) {
      const fileId = getFileIdByName(pic);
      if (fileId) {
        sheet.getRange(i + 2, 11).setValue(fileId);
        fixed++;
      }
    }
  });
  return { success: true, fixed: fixed };
}

function getDDImagesFolder() {
  const folders = DriveApp.getFoldersByName('DD Images');
  if (folders.hasNext()) return folders.next();
  return DriveApp.getRootFolder().createFolder('DD Images');
}

function getLogoBase64() {
  try {
    const file = DriveApp.getFileById('1He5NOU_Twr-CqctiFIuJFNe5wf-d-cTP');
    const blob = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    const mimeType = blob.getContentType();
    return { success: true, dataUrl: 'data:' + mimeType + ';base64,' + base64 };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function response(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
