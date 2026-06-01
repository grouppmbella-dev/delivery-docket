const SHEET_ID = '1w-vEijSeHI5rYtxjmtxyqv2Eh1P5yl8ShZHZlIftnOg';
const SHEET_NAME = 'DD';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    if (action === 'createDocket') return response(createDocket(data));
    if (action === 'updateSigned') return response(updateSigned(data));
    if (action === 'updateDocket') return response(updateDocket(data));
    if (action === 'deleteItem') return response(deleteItem(data));
    if (action === 'uploadImage') return response(uploadImage(data));
  } catch(err) {
    return response({ success: false, error: err.message });
  }
}

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'getDockets') return response(getDockets());
  if (action === 'getDocketById') return response(getDocketById(e.parameter.docNo));
  if (action === 'getLogoBase64') return response(getLogoBase64());
  if (action === 'fixPictureFileIds') return response(fixPictureFileIds());
}

function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
}

function getNextDocNo() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  const year = new Date().getFullYear();
  if (lastRow <= 1) return 'DDPM/' + year + '/0001';
  const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const docNos = data.map(r => r[0]).filter(v => v !== '');
  const unique = [...new Set(docNos)];
  const thisYear = unique.filter(d => String(d).includes('/' + year + '/'));
  if (thisYear.length === 0) return 'DDPM/' + year + '/0001';
  const nums = thisYear.map(d => parseInt(String(d).split('/')[2]));
  const next = Math.max(...nums) + 1;
  return 'DDPM/' + year + '/' + String(next).padStart(4, '0');
}

function getNextRowNo() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  // Filter only valid positive integers (not text, not #NUM!, not formulas)
  const nums = data
    .map(r => r[0])
    .filter(v => typeof v === 'number' && !isNaN(v) && isFinite(v) && v > 0 && Math.floor(v) === v);
  if (nums.length === 0) return 1;
  return Math.max(...nums) + 1;
}

function createDocket(data) {
  const sheet = getSheet();
  const docNo = getNextDocNo();
  const items = data.items;
  let rowNo = getNextRowNo();
  items.forEach((item, index) => {
    sheet.appendRow([
      rowNo++,
      docNo,
      data.requestDate,
      data.project,
      data.jobNumber,
      data.sendTo,
      index + 1,
      item.description,
      item.qty,
      item.unit,
      item.fileId || '',      // Store File ID for picture
      item.remark || '',
      data.nameIssueBy,
      data.issueDate,
      '', '',
      data.nameCheckBy,
      data.checkDate,
      '', '',
      data.nameDeliveryBy,
      data.deliveryDate,
      '', '',
      data.carNumberPlate,
      data.nameReceivedBy,
      data.receivedDate,
      '', '',
      'Draft'
    ]);
  });
  return { success: true, docNo: docNo };
}

function getDockets() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, data: [] };
  const data = sheet.getRange(2, 1, lastRow - 1, 30).getValues();
  const headers = ['docNoNo','docNo','requestDate','project','jobNumber','sendTo',
    'itemNo','description','qty','unit','picture','remark',
    'nameIssueBy','issueDate','issueSigned','issueSignedDate',
    'nameCheckBy','checkDate','checkSigned','checkSignedDate',
    'nameDeliveryBy','deliveryDate','deliverySigned','deliverySignedDate',
    'carNumberPlate','nameReceivedBy','receivedDate','receivedSigned',
    'receivedSignedDate','status'];
  const result = data.filter(row => row[1] !== '' && row[1] !== null).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
  return { success: true, data: result };
}

function getDocketById(docNo) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, 30).getValues();
  const headers = ['docNoNo','docNo','requestDate','project','jobNumber','sendTo',
    'itemNo','description','qty','unit','picture','remark',
    'nameIssueBy','issueDate','issueSigned','issueSignedDate',
    'nameCheckBy','checkDate','checkSigned','checkSignedDate',
    'nameDeliveryBy','deliveryDate','deliverySigned','deliverySignedDate',
    'carNumberPlate','nameReceivedBy','receivedDate','receivedSigned',
    'receivedSignedDate','status'];
  const items = data
    .filter(row => row[1] === docNo)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
  return { success: true, data: items };
}

function deleteItem(data) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  const allData = sheet.getRange(2, 1, lastRow - 1, 30).getValues();

  // Only allow delete on Draft dockets
  const docketRows = allData.map((row, i) => ({ row, i })).filter(r => r.row[1] === data.docNo);
  if (docketRows.length === 0) return { success: false, error: 'Docket not found' };
  if (docketRows[0].row[29] !== 'Draft') return { success: false, error: 'Only Draft dockets can be deleted' };

  // Find the specific item row to delete
  const targetRow = allData.findIndex((row, i) => 
    row[1] === data.docNo && String(row[6]) === String(data.itemNo)
  );
  if (targetRow === -1) return { success: false, error: 'Item not found' };

  // Delete the row
  sheet.deleteRow(targetRow + 2);

  // Check if docket has remaining items
  const remaining = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues()
    .filter(r => r[1] === data.docNo);

  // Renumber remaining items
  if (remaining.length > 0) {
    const newData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 30).getValues();
    let itemNum = 1;
    newData.forEach((row, i) => {
      if (row[1] === data.docNo) {
        sheet.getRange(i + 2, 7).setValue(itemNum++);
      }
    });
  }

  return { success: true, remainingItems: remaining.length };
}

function updateDocket(data) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  const allData = sheet.getRange(2, 1, lastRow - 1, 30).getValues();

  // Find all rows with this docNo
  const rowIndices = [];
  allData.forEach((row, i) => { if (row[1] === data.docNo) rowIndices.push(i); });
  if (rowIndices.length === 0) return { success: false, error: 'Docket not found' };

  const items = data.items;

  // Delete old rows (from bottom to top to avoid index shift)
  for (let i = rowIndices.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowIndices[i] + 2);
  }

  // Re-insert updated rows at original first position
  const insertAt = rowIndices[0] + 2;
  const existingFirst = allData[rowIndices[0]];

  items.forEach((item, index) => {
    sheet.insertRowBefore(insertAt + index);
    sheet.getRange(insertAt + index, 1, 1, 30).setValues([[
      existingFirst[0],          // Doc.no-no (keep original)
      data.docNo,                // Doc.no (keep)
      data.requestDate,
      data.project,
      data.jobNumber,
      data.sendTo,
      index + 1,
      item.description,
      item.qty,
      item.unit,
      item.fileId || item.picture || '',
      item.remark || '',
      data.nameIssueBy,
      data.issueDate,
      existingFirst[14] || '',   // issueSigned (keep)
      existingFirst[15] || '',   // issueSignedDate (keep)
      data.nameCheckBy,
      data.checkDate,
      existingFirst[18] || '',   // checkSigned (keep)
      existingFirst[19] || '',   // checkSignedDate (keep)
      data.nameDeliveryBy,
      data.deliveryDate,
      existingFirst[22] || '',   // deliverySigned (keep)
      existingFirst[23] || '',   // deliverySignedDate (keep)
      data.carNumberPlate,
      data.nameReceivedBy,
      data.receivedDate,
      existingFirst[27] || '',   // receivedSigned (keep)
      existingFirst[28] || '',   // receivedSignedDate (keep)
      existingFirst[29] || 'Draft' // status (keep)
    ]]);
  });

  return { success: true };
}

function updateSigned(data) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  const allData = sheet.getRange(2, 1, lastRow - 1, 30).getValues();
  const colMap = { 'issue': [15, 16], 'check': [19, 20], 'delivery': [23, 24], 'received': [29, 30] };
  const statusMap = { 'issue': 'Issued', 'delivery': 'Delivered', 'received': 'Completed' };
  const today = new Date().toLocaleDateString('en-GB');
  allData.forEach((row, index) => {
    if (row[1] === data.docNo) {
      const cols = colMap[data.type];
      sheet.getRange(index + 2, cols[0]).setValue('Signed');
      sheet.getRange(index + 2, cols[1]).setValue(today);
      if (statusMap[data.type]) sheet.getRange(index + 2, 30).setValue(statusMap[data.type]);
    }
  });
  return { success: true };
}

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

// Find File ID by filename in DD Images folder
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
  } catch(e) { return null; }
}

// Convert all filename-based pictures to File IDs in the sheet
function fixPictureFileIds() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 2) return { success: true, fixed: 0 };
  const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  let fixed = 0;
  data.forEach((row, i) => {
    const pic = String(row[10] || '').trim();
    // If picture looks like a filename (has extension) not a File ID
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

// Return logo as base64 so web app can embed directly (fixes CORS issue)
function getLogoBase64() {
  try {
    const file = DriveApp.getFileById('1He5NOU_Twr-CqctiFIuJFNe5wf-d-cTP');
    const blob = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    const mimeType = blob.getContentType();
    return { success: true, dataUrl: 'data:' + mimeType + ';base64,' + base64 };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function response(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
