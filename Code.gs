// Apps Script backend only: the frontend (index.html) is hosted on GitHub Pages.
// Do not use HtmlService/createHtmlOutputFromFile here.
const LOCK_WAIT_MS = 3000;
const COUNTED_TEXT = 'นับแล้ว';
const TIME_FORMAT = 'dd/MM/yyyy HH:mm:ss';

// ============================================================
// ตั้งค่าครั้งเดียว: แก้ค่าในฟังก์ชันนี้ -> เลือก setupConfig -> กด Run
// (ถ้าเคยรันไปแล้ว ไม่ต้องรันซ้ำ ค่าถูกเก็บใน Script Properties แล้ว)
// ============================================================
function setupConfig() {
  const SPREADSHEET_ID = ''; // เว้นว่างได้ ถ้าเปิด Apps Script จากในไฟล์ชีตนั้น
  const SHEET_NAMES = [      // ชื่อแท็บที่ต้องการให้ค้นหา ต้องตรงเป๊ะ
    'ครุภัณฑ์ อีสานกลาง',
    'วัสดุภัณฑ์ อีสานกลาง'
  ];

  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error('ไม่พบสเปรดชีต กรุณาใส่ SPREADSHEET_ID');
  }

  const existing = ss.getSheets().map(function (s) { return s.getName(); });
  SHEET_NAMES.forEach(function (name) {
    if (existing.indexOf(name) === -1) {
      throw new Error(
        'ไม่พบแท็บ "' + name + '" ชีตที่มี: ' +
        existing.map(function (n) { return '"' + n + '"'; }).join(', ')
      );
    }
  });

  PropertiesService.getScriptProperties().setProperties({
    SPREADSHEET_ID: ss.getId(),
    SHEET_NAMES: JSON.stringify(SHEET_NAMES)
  });

  Logger.log('บันทึกสำเร็จ: ไฟล์ "' + ss.getName() + '" แท็บ ' + SHEET_NAMES.join(', '));
}

// ============================================================

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = String(params.action || '').trim();
  const callback = String(params.callback || '').trim();

  let result;

  try {
    if (action === 'markAsCounted') {
      result = markAsCounted(params.code || '', params.rid || '');
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

// ---------- เปิดสเปรดชีตครั้งเดียวต่อ instance ----------
// ตัวแปร global จะค้างอยู่ตราบที่ instance ยังไม่หลับ
// คำขอถัดไปจึงไม่ต้องเปิดไฟล์/อ่าน Properties ซ้ำ (ลดเวลาได้เกือบวินาที)
let CTX_ = null;

function getCtx_() {
  if (CTX_) return CTX_;

  const props = PropertiesService.getScriptProperties().getProperties();
  if (!props.SPREADSHEET_ID || !props.SHEET_NAMES) {
    throw new Error('ยังไม่ได้ตั้งค่า กรุณารันฟังก์ชัน setupConfig ก่อน');
  }

  const ss = SpreadsheetApp.openById(props.SPREADSHEET_ID);
  const sheetNames = JSON.parse(props.SHEET_NAMES);

  const sheets = sheetNames.map(function (name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      const names = ss.getSheets().map(function (s) {
        return '"' + s.getName() + '"';
      }).join(', ');
      throw new Error(
        'ไม่พบชีต "' + name + '" ในไฟล์ "' + ss.getName() +
        '" (ชีตที่มี: ' + names + ')'
      );
    }
    return sheet;
  });

  CTX_ = {
    ss: ss,
    sheets: sheets,
    timezone: ss.getSpreadsheetTimeZone() || 'Asia/Bangkok'
  };
  return CTX_;
}

// ทำให้รหัสอยู่ในรูปมาตรฐานก่อนเทียบ เพื่อให้เทียบกันได้แม้
// - มีช่องว่าง/ขึ้นบรรทัดใหม่/ช่องว่างที่มองไม่เห็น  เช่น "ขก 50 01 08 005.ว" = "ขก500108005.ว"
// - ตัวพิมพ์เล็ก/ใหญ่ต่างกัน
// - เลขไทย (๐-๙) กับเลขอารบิก
// - อักษรไทยที่ประกอบสระ/วรรณยุกต์คนละรูปแบบ (Unicode NFC)
function normCode_(value) {
  return String(value == null ? '' : value)
    .normalize('NFC')
    .replace(/[\u0E50-\u0E59]/g, function (d) {
      return String(d.charCodeAt(0) - 0x0E50);
    })
    .replace(/[\s\u00A0\u200B-\u200D\u2060\uFEFF]+/g, '')
    .toLowerCase();
}

// ---------- ดัชนีรหัส (รหัสที่ทำให้เป็นมาตรฐานแล้ว -> ชีต/แถว) ----------
// เดิมทุกครั้งที่สแกนจะ normalize รหัสของ "ทุกแถว" ซึ่งรหัสภาษาไทยช้ากว่า ASCII มาก
// ตอนนี้อ่านเฉพาะคอลัมน์ B แล้วสร้างดัชนีครั้งเดียวต่อ instance
// สแกนครั้งต่อ ๆ ไปจึงเป็นการ lookup ทันที + อ่านแถวที่เจอแถวเดียว
function buildIndex_(ctx) {
  const map = Object.create(null);
  let hasData = false;

  for (let i = 0; i < ctx.sheets.length; i++) {
    const sh = ctx.sheets[i];
    const lastRow = sh.getLastRow();
    if (lastRow < 2) continue;
    hasData = true;

    const col = sh.getRange(2, 2, lastRow - 1, 1).getValues();
    for (let r = 0; r < col.length; r++) {
      const key = normCode_(col[r][0]);
      if (key && !(key in map)) map[key] = { s: i, r: r + 2 }; // เจอก่อนใช้ก่อน เหมือนเดิม
    }
  }

  ctx.index = { map: map, hasData: hasData };
  return { map: map, hasData: hasData, justBuilt: true };
}

function getIndex_(ctx, force) {
  if (!force && ctx.index) {
    return { map: ctx.index.map, hasData: ctx.index.hasData, justBuilt: false };
  }
  return buildIndex_(ctx);
}

function findAsset_(ctx, code) {
  const target = normCode_(code);
  let hasData = false;

  for (let pass = 0; pass < 2; pass++) {
    const idx = getIndex_(ctx, pass === 1); // รอบสอง = สร้างดัชนีใหม่ (กรณีมีแถวใหม่/แก้รหัสในชีต)
    hasData = idx.hasData;

    const hit = idx.map[target];
    if (hit) {
      const sh = ctx.sheets[hit.s];
      const values = sh.getRange(hit.r, 1, 1, 9).getValues()[0];
      // ยืนยันว่าแถวนั้นยังเป็นรหัสเดิมจริง (กันดัชนีเก่า)
      if (normCode_(values[1]) === target) {
        return { hasData: true, sheet: sh, row: hit.r, values: values };
      }
    }
    if (idx.justBuilt) break; // เพิ่งสร้างใหม่แล้วไม่เจอ = ไม่มีจริง
  }

  return { hasData: hasData, sheet: null, row: 0, values: null };
}

function markAsCounted(rawCode, rid) {
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
    // ถ้าเป็นคำขอซ้ำ (หน้าเว็บส่งซ้ำอัตโนมัติเพราะรอนาน) ให้ตอบผลเดิม
    // ไม่ใช่ "นับแล้ว" ทั้งที่เพิ่งนับสำเร็จจากคำขอแรก
    const ridKey = rid ? 'rid_' + String(rid).slice(0, 80) : '';
    if (ridKey) {
      try {
        const prev = CacheService.getScriptCache().get(ridKey);
        if (prev) return JSON.parse(prev);
      } catch (_) {}
    }

    const ctx = getCtx_();
    const found = findAsset_(ctx, code);

    if (!found.hasData) {
      return {
        ok: false,
        type: 'EMPTY_SHEET',
        message: 'ยังไม่มีข้อมูลครุภัณฑ์ในชีต'
      };
    }

    if (!found.sheet) {
      return {
        ok: false,
        type: 'NOT_FOUND',
        message: 'ไม่พบข้อมูล',
        code: code
      };
    }

    const sheet = found.sheet;
    const row = found.row;
    const values = found.values;

    const asset = {
      sheet: sheet.getName(),
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

    if (asset.counted === COUNTED_TEXT || asset.countedAt) {
      return {
        ok: true,
        type: 'ALREADY_COUNTED',
        message: 'รายการนี้นับแล้ว',
        asset: serializeAsset_(asset, ctx.timezone)
      };
    }

    const now = new Date();

    sheet.getRange(row, 6, 1, 2).setValues([[COUNTED_TEXT, now]]);
    sheet.getRange(row, 7).setNumberFormat(TIME_FORMAT);

    asset.counted = COUNTED_TEXT;
    asset.countedAt = now;

    // ไม่ต้อง flush(): ข้อมูลถูกบันทึกเมื่อจบการทำงานอยู่แล้ว

    const done = {
      ok: true,
      type: 'COUNTED',
      message: 'นับรายการเรียบร้อย',
      asset: serializeAsset_(asset, ctx.timezone)
    };
    if (ridKey) {
      try {
        CacheService.getScriptCache().put(ridKey, JSON.stringify(done), 600);
      } catch (_) {}
    }
    return done;
  } finally {
    lock.releaseLock();
  }
}

function serializeAsset_(asset, timezone) {
  return {
    sheet: asset.sheet,
    row: asset.row,
    sequence: asset.sequence,
    code: asset.code,
    name: asset.name,
    amount: asset.amount,
    currentQuantity: asset.currentQuantity,
    counted: asset.counted,
    countedAt: asset.countedAt
      ? Utilities.formatDate(new Date(asset.countedAt), timezone, TIME_FORMAT)
      : '',
    location: asset.location,
    note: asset.note
  };
}

function ping_() {
  const ctx = getCtx_();
  getIndex_(ctx, false); // อุ่นดัชนีไว้ล่วงหน้า สแกนจริงจะได้ไม่ต้องรอ

  return {
    ok: true,
    type: 'PING',
    message: 'เชื่อมต่อ Apps Script และชีตสำเร็จ',
    sheets: ctx.sheets.map(function (sh) {
      return { name: sh.getName(), lastRow: sh.getLastRow() };
    })
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