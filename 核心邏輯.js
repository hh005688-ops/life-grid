/**
 * =================================================================
 * 四維系統 (Four-Dimensional System) - 十年高效查表萬年曆核心
 * =================================================================
 * 讀寫原則：試算表一律一次 getValues() / setValues() 進記憶體。
 * 迴圈只處理記憶體陣列，不再對單一儲存格呼叫 getValue() / setValue()。
 */

/**
 * 瀏覽器直接開啟 /exec：回傳網頁。
 * 帶 ?action= 的請求（Vercel 的 fetch）：回傳 JSON。
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return jsonResponse_(dispatchApi_(e.parameter.action, e.parameter));
  }
  try {
    return HtmlService.createTemplateFromFile('index')
        .evaluate()
        .setTitle('四維系統 - 智慧排班與記帳')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return HtmlService.createTemplateFromFile('前端架構')
        .evaluate()
        .setTitle('四維系統 - 智慧排班與記帳')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

/**
 * Vercel 寫入用。前端以 text/plain 送 JSON，避免瀏覽器先發 OPTIONS。
 * GAS Web App 不會回應預檢請求。
 */
function doPost(e) {
  var data = {};
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return jsonResponse_({ error: '請求內容不是 JSON' });
  }
  if (e && e.parameter) {
    Object.keys(e.parameter).forEach(function (key) {
      if (data[key] === undefined) data[key] = e.parameter[key];
    });
  }
  return jsonResponse_(dispatchApi_(data.action, data));
}

/**
 * 密鑰存放在「專案設定 → 指令碼屬性」的 API_KEY，不放進程式碼，避免被推上 GitHub。
 * 第一次使用：在編輯器執行 generateApiKey()，再從執行記錄複製密鑰。
 * 未設定 API_KEY 時一律拒絕（fail closed）。
 */
var API_KEY_PROPERTY = 'API_KEY';

function generateApiKey() {
  const key = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty(API_KEY_PROPERTY, key);
  Logger.log('新的 API 密鑰（請貼到 Vercel 頁面的「API 密鑰」欄位）：' + key);
  return key;
}

function isAuthorized_(provided) {
  const expected = PropertiesService.getScriptProperties().getProperty(API_KEY_PROPERTY);
  if (!expected) return false;
  return safeEquals_(String(provided || ''), expected);
}

// 逐字比對且不提早結束，避免從回應時間猜出密鑰
function safeEquals_(a, b) {
  let diff = a.length ^ b.length;
  const length = Math.max(a.length, b.length);
  for (let i = 0; i < length; i++) {
    diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  }
  return diff === 0;
}

function dispatchApi_(action, data) {
  var payload = data || {};
  console.log('dispatchApi ' + action);
  if (!isAuthorized_(payload.key)) {
    const configured = !!PropertiesService.getScriptProperties().getProperty(API_KEY_PROPERTY);
    console.warn('dispatchApi unauthorized ' + action);
    return {
      code: 'UNAUTHORIZED',
      error: configured
        ? 'API 密鑰錯誤或未填寫'
        : '後端尚未設定密鑰，請先在 Apps Script 執行 generateApiKey()'
    };
  }
  try {
    switch (action) {
      case 'getCalendarMonthData':
        return getCalendarMonthData(payload.sheetName, payload.year, payload.month);
      case 'getFixedSettings':
        return getFixedSettings();
      case 'saveFixedSettings':
        return saveFixedSettings(payload.baseDate, payload.patternText);
      case 'updateRecordFields':
        return updateRecordFields(payload.sheetName, payload.dateStr, parseFields_(payload.fields));
      default:
        return { error: '未知的 action：' + action };
    }
  } catch (error) {
    return { error: error.message };
  }
}

function parseFields_(fields) {
  if (!fields) return {};
  if (typeof fields === 'string') {
    try {
      return JSON.parse(fields);
    } catch (e) {
      return {};
    }
  }
  return fields;
}

function jsonResponse_(payload) {
  return ContentService
      .createTextOutput(JSON.stringify(payload))
      .setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 回傳指定月份的班表。另附 timings，讓瀏覽器主控台能對照伺服器耗時。
 * 固定輪班會順便帶上基準日與樣板，避免前端再打一支 getFixedSettings。
 */
function getCalendarMonthData(sheetName, year, month) {
  const timings = {};
  const totalStart = Date.now();
  console.time('getCalendarMonthData');

  try {
    year = Number(year);
    month = Number(month);
    if (!isFinite(year) || !isFinite(month)) {
      throw new Error('年份或月份無效');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = sheetName || '隨機排班';
    const sheet = ss.getSheetByName(targetSheet);
    if (!sheet) {
      throw new Error('找不到「' + targetSheet + '」分頁，請確認分頁名稱是否正確。');
    }

    const snapshot = measure_('readSheetBundle', timings, 'readSheetMs', function () {
      return readSheetSnapshot_(sheet);
    });
    const holidayMap = measure_('getTenYearHolidayCache', timings, 'holidayMs', function () {
      return getTenYearHolidayCache(year);
    });
    const records = measure_('buildMonthRecords', timings, 'buildMonthMs', function () {
      return buildMonthRecords_(snapshot, holidayMap, targetSheet, year, month);
    });

    timings.sheetRows = snapshot.rowCount;
    timings.recordCount = records.length;

    const payload = {
      records: records,
      timings: timings
    };
    if (targetSheet === '固定輪班') {
      payload.fixedSettings = {
        success: true,
        baseDate: formatDateKey(snapshot.baseDate),
        patternText: snapshot.pattern.join(', ')
      };
    }
    return payload;
  } catch (error) {
    Logger.log('動態萬年曆生成錯誤: ' + error.message);
    return { error: error.message, timings: timings };
  } finally {
    timings.totalMs = Date.now() - totalStart;
    console.timeEnd('getCalendarMonthData');
    console.log('getCalendarMonthData timings ' + JSON.stringify(timings));
  }
}

/**
 * A:I 一次讀進記憶體。A-F 是每日紀錄，G2 是起始日，H2 以下是循環樣板，I2 是循環天數。
 */
function readSheetSnapshot_(sheet) {
  const lastRow = sheet.getLastRow();
  const snapshot = {
    rowCount: lastRow,
    sheetDataMap: {},
    baseDate: defaultBaseDate_(),
    pattern: defaultPattern_()
  };
  if (lastRow < 1) return snapshot;

  const rows = sheet.getRange(1, 1, lastRow, 9).getValues();
  const config = parseShiftConfig_(rows);
  snapshot.baseDate = config.baseDate;
  snapshot.pattern = config.pattern;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const dateStr = formatDateKey(row[0]);
    if (!dateStr) continue;
    snapshot.sheetDataMap[dateStr] = {
      row: i + 1,
      manualShift: String(row[2] || '').trim(),
      category: String(row[3] || '').trim(),
      amount: Number(row[4]) || 0,
      memo: String(row[5] || '').trim()
    };
  }
  return snapshot;
}

function readShiftConfig_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    return { baseDate: defaultBaseDate_(), pattern: defaultPattern_() };
  }
  const rows = sheet.getRange(1, 1, lastRow, 9).getValues();
  return parseShiftConfig_(rows);
}

/**
 * 起始日在 G2（rows[1][6]）；循環樣板從 H2（rows[1][7]）往下讀到空值為止。
 */
function parseShiftConfig_(rows) {
  const baseDate = parseBaseDate_(rows.length > 1 ? rows[1][6] : '');
  const customPattern = [];
  for (let i = 1; i < rows.length; i++) {
    const val = rows[i][7];
    if (val === '' || val === null || val === undefined) break;
    const text = String(val).trim();
    if (text === '') break;
    customPattern.push(text);
  }
  return {
    baseDate: baseDate,
    pattern: customPattern.length ? customPattern : defaultPattern_()
  };
}

function parseBaseDate_(h1Val) {
  if (h1Val instanceof Date && !isNaN(h1Val.getTime())) return h1Val;
  if (h1Val) {
    const parsed = new Date(h1Val);
    if (!isNaN(parsed.getTime())) return parsed;
  }
  return defaultBaseDate_();
}

function buildMonthRecords_(snapshot, holidayMap, targetSheet, year, month) {
  const totalDays = new Date(year, month + 1, 0).getDate();
  const weekDays = ['週日', '週一', '週二', '週三', '週四', '週五', '週六'];
  const pattern = snapshot.pattern;
  const patternLength = pattern.length || 1;
  const baseDate = snapshot.baseDate;
  const baseTime = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate()).getTime();
  const result = [];

  for (let day = 1; day <= totalDays; day++) {
    const current = new Date(year, month, day);
    const dateStr = formatDateKey(current);
    const targetTime = new Date(current.getFullYear(), current.getMonth(), current.getDate()).getTime();
    let diffDays = Math.round((targetTime - baseTime) / 86400000);
    let patternIndex = diffDays % patternLength;
    if (patternIndex < 0) {
      patternIndex = (patternIndex % patternLength + patternLength) % patternLength;
    }

    const sheetRecord = snapshot.sheetDataMap[dateStr] || {};

    // 班別：手動值優先；沒有時固定輪班用樣板，隨機排班為空
    let finalShift = '';
    if (sheetRecord.manualShift) {
      finalShift = sheetRecord.manualShift;
    } else if (targetSheet === '固定輪班') {
      finalShift = pattern[patternIndex] || '';
    }

    // 舊資料可能把節日名稱誤植在 D 欄：非收支類別者視為節日
    let category = sheetRecord.category || '';
    let extractedHolidayFromCat = '';
    if (category && !isExpenseCategory_(category)) {
      extractedHolidayFromCat = category;
      category = '';
    }

    const holiday = shortenHolidayName_(holidayMap[dateStr] || extractedHolidayFromCat || '');
    let memo = sheetRecord.memo || '';
    if (holiday && memo.indexOf(holiday) === -1) {
      memo = '[' + holiday + ']' + (memo ? ' ' + memo : '');
    }

    result.push({
      row: sheetRecord.row || null,
      date: dateStr,
      weekDay: weekDays[current.getDay()],
      shift: finalShift,
      category: category,
      amount: sheetRecord.amount || 0,
      memo: memo,
      holiday: holiday
    });
  }
  return result;
}

var EXPENSE_CATEGORY_KEYWORDS_ = ['餐飲', '飲食', '交通', '油資', '加油', '保養', '加班', '獎金', '薪資'];

function isExpenseCategory_(category) {
  const text = String(category || '');
  for (let i = 0; i < EXPENSE_CATEGORY_KEYWORDS_.length; i++) {
    if (text.indexOf(EXPENSE_CATEGORY_KEYWORDS_[i]) !== -1) return true;
  }
  return false;
}

var HOLIDAY_NAME_MAP_ = {
  '中華民國開國紀念日': '元旦',
  '和平紀念日': '228紀念日',
  '兒童節及民族掃墓節': '清明節',
  '民族掃墓節': '清明節',
  '勞動節': '勞動節',
  '孔子誕辰紀念日': '教師節',
  '孔子誕辰紀念日（教師節）': '教師節',
  '中華民國國慶日': '國慶日',
  '中華民國國慶日 補假': '國慶補假',
  '中華民國國慶日補假': '國慶補假',
  '台灣光復節': '光復節',
  '臺灣光復暨金門古寧頭大捷紀念日': '光復節',
  '台灣光復節 補假': '光復節補假',
  '台灣光復節補假': '光復節補假',
  '行憲紀念日': '行憲紀念日',
  // 已簡化的名稱對應自身，讓函式可重複套用（快取讀取時也會再轉一次）
  '元旦': '元旦',
  '228紀念日': '228紀念日',
  '清明節': '清明節',
  '教師節': '教師節',
  '國慶日': '國慶日',
  '國慶補假': '國慶補假',
  '光復節': '光復節',
  '光復節補假': '光復節補假'
};

/**
 * 國定假日簡稱，結果一律在 6 字以內。
 */
function shortenHolidayName_(rawTitle) {
  let text = String(rawTitle || '').trim();
  if (!text) return '';
  if (HOLIDAY_NAME_MAP_.hasOwnProperty(text)) return HOLIDAY_NAME_MAP_[text];
  const compact = text.replace(/\s+/g, '');
  if (HOLIDAY_NAME_MAP_.hasOwnProperty(compact)) return HOLIDAY_NAME_MAP_[compact];

  const hasMakeup = compact.indexOf('補假') !== -1;
  text = compact.replace(/補假/g, '').replace(/中華民國|紀念日|（|）|\(|\)/g, '');
  if (hasMakeup) text += '補假';
  if (text.length > 6) text = text.slice(0, 6);
  return text;
}

/**
 * 五年一段的國定假日快取。命中時不再連線 Google 日曆。
 */
function getTenYearHolidayCache(targetYear) {
  const blockYear = Math.floor(targetYear / 5) * 5;
  const cache = CacheService.getScriptCache();
  const cacheKey = 'holidays_block_v2_' + blockYear;

  console.time('holidayCacheLookup');
  let cachedData = null;
  try {
    cachedData = cache.get(cacheKey);
  } finally {
    console.timeEnd('holidayCacheLookup');
  }

  if (cachedData) {
    try {
      console.log('holiday cache hit ' + cacheKey);
      const cachedMap = JSON.parse(cachedData);
      Object.keys(cachedMap).forEach(function (key) {
        cachedMap[key] = shortenHolidayName_(cachedMap[key]);
      });
      return cachedMap;
    } catch (e) {
      Logger.log('假日快取解析失敗，改向日曆重抓: ' + e.message);
    }
  }

  console.log('holiday cache miss ' + cacheKey);
  const holidayMap = {};
  console.time('holidayCalendarFetch');
  try {
    const calendarId = 'zh-tw.taiwan#holiday@group.v.calendar.google.com';
    const cal = CalendarApp.getCalendarById(calendarId);
    if (cal) {
      const startDate = new Date(blockYear, 0, 1);
      const endDate = new Date(blockYear + 5, 11, 31);
      const events = cal.getEvents(startDate, endDate);
      for (let i = 0; i < events.length; i++) {
        const event = events[i];
        holidayMap[formatDateKey(event.getStartTime())] = shortenHolidayName_(event.getTitle());
      }
    }
    cache.put(cacheKey, JSON.stringify(holidayMap), 86400);
  } catch (err) {
    Logger.log('抓取十年官方行事曆例外: ' + err.message);
  } finally {
    console.timeEnd('holidayCalendarFetch');
  }
  return holidayMap;
}

/**
 * 一次寫回班別、類別、金額、備註。前端存檔只打這一支。
 */
function updateRecordFields(sheetName, dateStr, fields) {
  const timings = {};
  const started = Date.now();
  console.time('updateRecordFields');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName || '隨機排班');
    if (!sheet) throw new Error('找不到對應分頁');

    const normalized = normalizeFields_(fields);
    let targetRow = findDateRow_(sheet, dateStr);
    if (targetRow === -1) {
      const parsed = new Date(dateStr);
      const weekDays = ['週日', '週一', '週二', '週三', '週四', '週五', '週六'];
      targetRow = sheet.getLastRow() + 1;
      sheet.getRange(targetRow, 1, 1, 6).setValues([[
        dateStr,
        weekDays[parsed.getDay()],
        normalized.shift,
        normalized.category,
        normalized.amount,
        normalized.memo
      ]]);
    } else {
      sheet.getRange(targetRow, 3, 1, 4).setValues([[
        normalized.shift,
        normalized.category,
        normalized.amount,
        normalized.memo
      ]]);
    }
    return { success: true, timings: timings };
  } catch (error) {
    return { success: false, error: error.message, timings: timings };
  } finally {
    timings.totalMs = Date.now() - started;
    console.timeEnd('updateRecordFields');
    console.log('updateRecordFields timings ' + JSON.stringify(timings));
  }
}

/**
 * 保留單欄更新，供舊呼叫使用。寫入仍是一次 setValues，不在迴圈裡寫儲存格。
 */
function updateRecordByDate(sheetName, dateStr, fieldName, value) {
  const timings = {};
  const started = Date.now();
  console.time('updateRecordByDate');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName || '隨機排班');
    if (!sheet) throw new Error('找不到對應分頁');

    const fieldMap = { shift: 3, category: 4, amount: 5, memo: 6 };
    const colIndex = fieldMap[fieldName];
    if (!colIndex) throw new Error('無效的欄位名稱');

    let targetRow = findDateRow_(sheet, dateStr);
    if (targetRow === -1) {
      const parsed = new Date(dateStr);
      const weekDays = ['週日', '週一', '週二', '週三', '週四', '週五', '週六'];
      targetRow = sheet.getLastRow() + 1;
      sheet.getRange(targetRow, 1, 1, 6).setValues([[
        dateStr,
        weekDays[parsed.getDay()],
        '',
        '',
        0,
        ''
      ]]);
    }

    let finalVal = value;
    if (fieldName === 'amount') finalVal = Number(value) || 0;
    sheet.getRange(targetRow, colIndex, 1, 1).setValues([[finalVal]]);
    return { success: true, timings: timings };
  } catch (error) {
    return { success: false, error: error.message, timings: timings };
  } finally {
    timings.totalMs = Date.now() - started;
    console.timeEnd('updateRecordByDate');
    console.log('updateRecordByDate timings ' + JSON.stringify(timings));
  }
}

function findDateRow_(sheet, dateStr) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < dates.length; i++) {
    if (formatDateKey(dates[i][0]) === dateStr) return i + 2;
  }
  return -1;
}

function normalizeFields_(fields) {
  const source = fields || {};
  return {
    shift: source.shift != null ? String(source.shift) : '',
    category: source.category != null ? String(source.category) : '',
    amount: Number(source.amount) || 0,
    memo: source.memo != null ? String(source.memo) : ''
  };
}

function getFixedSettings() {
  const timings = {};
  const started = Date.now();
  console.time('getFixedSettings');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('固定輪班');
    if (!sheet) throw new Error('找不到「固定輪班」分頁');

    const config = measure_('readShiftConfig', timings, 'readConfigMs', function () {
      return readShiftConfig_(sheet);
    });
    return {
      success: true,
      baseDate: formatDateKey(config.baseDate),
      patternText: config.pattern.join(', '),
      timings: timings
    };
  } catch (error) {
    return { success: false, error: error.message, timings: timings };
  } finally {
    timings.totalMs = Date.now() - started;
    console.timeEnd('getFixedSettings');
    console.log('getFixedSettings timings ' + JSON.stringify(timings));
  }
}

function saveFixedSettings(baseDate, patternText) {
  const timings = {};
  const started = Date.now();
  console.time('saveFixedSettings');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('固定輪班');
    if (!sheet) throw new Error('找不到「固定輪班」分頁');

    const patternArray = String(patternText || '')
      .split(/[,，、\s\n]+/)
      .map(function (item) { return item.trim(); })
      .filter(function (item) { return item !== ''; });

    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      sheet.getRange(2, 8, lastRow - 1, 1).clearContent();
    }
    if (patternArray.length > 0) {
      const writeData = patternArray.map(function (item) { return [item]; });
      sheet.getRange(2, 8, writeData.length, 1).setValues(writeData);
      sheet.getRange(2, 9, 1, 1).setValues([[writeData.length]]);
    }
    sheet.getRange(2, 7, 1, 1).setValues([[baseDate]]);

    return { success: true, timings: timings };
  } catch (error) {
    return { success: false, error: error.message, timings: timings };
  } finally {
    timings.totalMs = Date.now() - started;
    console.timeEnd('saveFixedSettings');
    console.log('saveFixedSettings timings ' + JSON.stringify(timings));
  }
}

function formatDateKey(dateInput) {
  if (!dateInput) return '';
  let parsed = new Date(dateInput);
  if (isNaN(parsed.getTime())) {
    parsed = new Date(String(dateInput).trim().replace(/\//g, '-'));
  }
  if (isNaN(parsed.getTime())) return '';

  const year = parsed.getFullYear();
  const month = String(parsed.getMonth() + 1).padStart(2, '0');
  const day = String(parsed.getDate()).padStart(2, '0');
  return year + '-' + month + '-' + day;
}

function defaultBaseDate_() {
  return new Date(2026, 8, 1);
}

function defaultPattern_() {
  return ['晚班', '晚班', '晚班', '晚班', '休假', '休假', '早班', '早班', '早班', '早班', '休假', '休假'];
}

function measure_(label, timings, field, fn) {
  console.time(label);
  const started = Date.now();
  try {
    return fn();
  } finally {
    console.timeEnd(label);
    timings[field] = Date.now() - started;
  }
}
