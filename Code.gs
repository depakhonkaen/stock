const SHEET_NAME = 'Stock';

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = String(params.action || '').trim();

  // GitHub Pages เรียก Apps Script ผ่าน JSONP
  // เพื่อให้กล้องอยู่บน GitHub Pages แต่ข้อมูลยังบันทึกใน Google Sheet
  if (action === 'markAsCounted') {
    let result;
    try {
      result = markAsCounted(params.code || '');
    } catch (err) {
      result = {
        ok: false,
        type: 'INVALID',
        message: String((err && err.message) || err)
      };
    }

    const callback = String(params.callback || '').trim();

    if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
      return ContentService
        .createTextOutput(callback + '(' + JSON.stringify(result) + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // เปิดหน้า Apps Script โดยตรงยังใช้งานได้
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบตรวจนับครุภัณฑ์')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('ไม่พบชีต "' + SHEET_NAME + '"');
  return sheet;
}

function markAsCounted(rawCode) {
  const code = String(rawCode || '').trim();
  if (!code) return { ok: false, type: 'INVALID', message: 'กรุณาระบุรหัสครุภัณฑ์' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getSheet_();
    const rowCount = Math.max(sheet.getLastRow() - 1, 1);
    const cell = sheet.getRange(2, 2, rowCount, 1)
      .createTextFinder(code)
      .matchEntireCell(true)
      .matchCase(false)
      .findNext();

    if (!cell) {
      return { ok: false, type: 'NOT_FOUND', message: 'ไม่พบข้อมูล', code };
    }

    const row = cell.getRow();
    const values = sheet.getRange(row, 1, 1, 9).getValues()[0];

    const asset = {
      row,
      sequence: values[0],
      code: String(values[1] ?? ''),
      name: String(values[2] ?? ''),
      amount: values[3],
      currentQuantity: values[4],
      counted: String(values[5] ?? '').trim(),
      countedAt: values[6],
      location: String(values[7] ?? ''),
      note: String(values[8] ?? '')
    };

    if (asset.counted === 'นับแล้ว' || asset.countedAt) {
      return { ok: true, type: 'ALREADY_COUNTED', message: 'รายการนี้นับแล้ว', asset: serializeAsset_(asset) };
    }

    const now = new Date();
    sheet.getRange(row, 6, 1, 2).setValues([['นับแล้ว', now]]);
    sheet.getRange(row, 7).setNumberFormat('dd/MM/yyyy HH:mm:ss');

    asset.counted = 'นับแล้ว';
    asset.countedAt = now;

    return { ok: true, type: 'COUNTED', message: 'นับรายการเรียบร้อย', asset: serializeAsset_(asset) };
  } finally {
    lock.releaseLock();
  }
}

function serializeAsset_(asset) {
  const timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'Asia/Bangkok';
  return {
    row: asset.row,
    sequence: asset.sequence,
    code: asset.code,
    name: asset.name,
    amount: asset.amount,
    currentQuantity: asset.currentQuantity,
    counted: asset.counted,
    countedAt: asset.countedAt
      ? Utilities.formatDate(new Date(asset.countedAt), timezone, 'dd/MM/yyyy HH:mm:ss')
      : '',
    location: asset.location,
    note: asset.note
  };
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}