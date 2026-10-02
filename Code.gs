const SHEET_NAME = 'Stock';
const LOCK_WAIT_MS = 3000;

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = String(params.action || '').trim();
  const callback = String(params.callback || '').trim();

  let result;

  try {
    if (action === 'markAsCounted') {
      result = markAsCounted(params.code || '');
    } else if (action === 'ping') {
      result = ping_();
    } else {
      result = {
        ok: true,
        type: 'READY',
        message: 'Stock API พร้อมใช้งาน'
      };
    }
  } catch (err) {
    result = {
      ok: false,
      type: 'SERVER_ERROR',
      message: getErrorMessage_(err),
      errorName: String((err && err.name) || 'Error')
    };
  }

  return output_(result, callback);
}

function output_(result, callback) {
  const json = JSON.stringify(result);

  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('ไม่พบ Spreadsheet ที่ผูกกับ Apps Script นี้');
  }

  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('ไม่พบชีต "' + SHEET_NAME + '"');
  }

  return sheet;
}

function markAsCounted(rawCode) {
  const code = String(rawCode || '').trim();

  if (!code) {
    return {
      ok: false,
      type: 'INVALID',
      message: 'กรุณาระบุรหัสครุภัณฑ์'
    };
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(LOCK_WAIT_MS);
  } catch (err) {
    return {
      ok: false,
      type: 'LOCK_TIMEOUT',
      message: 'ระบบกำลังประมวลผลรายการอื่นอยู่ กรุณาลองสแกนอีกครั้ง',
      detail: getErrorMessage_(err)
    };
  }

  try {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return {
        ok: false,
        type: 'EMPTY_SHEET',
        message: 'ชีต Stock ยังไม่มีข้อมูลครุภัณฑ์'
      };
    }

    const rowCount = lastRow - 1;
    const cell = sheet.getRange(2, 2, rowCount, 1)
      .createTextFinder(code)
      .matchEntireCell(true)
      .matchCase(false)
      .findNext();

    if (!cell) {
      return {
        ok: false,
        type: 'NOT_FOUND',
        message: 'ไม่พบข้อมูล',
        code: code
      };
    }

    const row = cell.getRow();
    const values = sheet.getRange(row, 1, 1, 9).getValues()[0];

    const asset = {
      row: row,
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
      return {
        ok: true,
        type: 'ALREADY_COUNTED',
        message: 'รายการนี้นับแล้ว',
        asset: serializeAsset_(asset)
      };
    }

    const now = new Date();

    sheet.getRange(row, 6, 1, 2).setValues([
      ['นับแล้ว', now]
    ]);
    sheet.getRange(row, 7).setNumberFormat('dd/MM/yyyy HH:mm:ss');

    asset.counted = 'นับแล้ว';
    asset.countedAt = now;

    SpreadsheetApp.flush();

    return {
      ok: true,
      type: 'COUNTED',
      message: 'นับรายการเรียบร้อย',
      asset: serializeAsset_(asset)
    };
  } finally {
    lock.releaseLock();
  }
}

function serializeAsset_(asset) {
  const timezone =
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() ||
    'Asia/Bangkok';

  return {
    row: asset.row,
    sequence: asset.sequence,
    code: asset.code,
    name: asset.name,
    amount: asset.amount,
    currentQuantity: asset.currentQuantity,
    counted: asset.counted,
    countedAt: asset.countedAt
      ? Utilities.formatDate(
          new Date(asset.countedAt),
          timezone,
          'dd/MM/yyyy HH:mm:ss'
        )
      : '',
    location: asset.location,
    note: asset.note
  };
}

function ping_() {
  const sheet = getSheet_();

  return {
    ok: true,
    type: 'PING',
    message: 'เชื่อมต่อ Apps Script และชีต Stock สำเร็จ',
    sheet: SHEET_NAME,
    lastRow: sheet.getLastRow()
  };
}

function getErrorMessage_(err) {
  if (!err) return 'ไม่ทราบสาเหตุ';

  const message = String(err.message || err || '').trim();
  if (message) return message;

  try {
    return JSON.stringify(err);
  } catch (_) {
    return String(err);
  }
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}
