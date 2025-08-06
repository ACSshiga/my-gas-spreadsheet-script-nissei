/**
 * Utils.gs
 * 汎用ユーティリティ関数を定義します。
 */

// =================================================================================
// === 日付関連ユーティリティ ===
// =================================================================================
function formatDateForComparison(date) {
  if (!isValidDate(date)) return null;
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function getJapaneseHolidays(year) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `holidays_${year}`;
  const cached = cache.get(cacheKey);
  if (cached) {
    return new Set(JSON.parse(cached));
  }
  try {
    const calendar = CalendarApp.getCalendarById(CONFIG.HOLIDAY_CALENDAR_ID);
    if (!calendar) {
      console.warn("祝日カレンダーが見つかりません。");
      return new Set();
    }
    const startDate = new Date(year, 0, 1);
    const endDate = new Date(year, 11, 31);
    const events = calendar.getEvents(startDate, endDate);
    const holidays = new Set(events.map(e => formatDateForComparison(e.getStartTime())));
    cache.put(cacheKey, JSON.stringify([...holidays]), 21600);
    return holidays;
  } catch (e) {
    console.error("祝日取得エラー:", e);
    return new Set();
  }
}

function isHoliday(date, holidaySet) {
  if (!isValidDate(date)) return false;
  const dayOfWeek = date.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) return true;
  return holidaySet.has(formatDateForComparison(date));
}

function isValidDate(value) {
  return value instanceof Date && !isNaN(value.getTime());
}

// =================================================================================
// === 文字列・データ処理ユーティリティ ===
// =================================================================================
function safeTrim(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function createHyperlinkFormula(url, displayText) {
  if (!url) return "";
  const safeUrl = String(url).replace(/"/g, '""');
  const safeText = String(displayText || url).replace(/"/g, '""');
  return `=HYPERLINK("${safeUrl}", "${safeText}")`;
}

function toNumber(value) {
  const num = parseFloat(value);
  return isNaN(num) ? 0 : num;
}

// =================================================================================
// === マスタデータ取得ユーティリティ ===
// =================================================================================
/**
 * マスタシートからデータを取得します。
 * @param {string} masterSheetName - マスタシート名
 * @param {number} numColumns - 取得する列数 (省略可能)
 */
function getMasterData(masterSheetName, numColumns) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(masterSheetName);
  if (!sheet) return [];

  const colsToFetch = numColumns || sheet.getLastColumn();

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, colsToFetch).getValues();
  return values.filter(row => row[0] !== "");
}

/**
 * 各マスタシートから色の対応表（Map）を作成します。
 * @param {string} sheetName - マスタシート名
 * @param {number} keyColIndex - キーとなる項目が含まれる列（0始まり）
 * @param {number} colorColIndex - カラーコードが含まれる列（0始まり）
 */
function getColorMapFromMaster(sheetName, keyColIndex, colorColIndex) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `color_map_${sheetName}`;
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      return new Map(JSON.parse(cached));
    } catch (e) { /* ignore */ }
  }

  const masterData = getMasterData(sheetName);
  const colorMap = new Map(
    masterData.map(row => [row[keyColIndex], row[colorColIndex]])
  );

  cache.put(cacheKey, JSON.stringify([...colorMap]), 3600); // 1時間キャッシュ
  return colorMap;
}


function getTantoushaNameByEmail(email) {
  if (!email) return null;
  const userEmail = email.trim();
  const masterData = getMasterData(CONFIG.SHEETS.TANTOUSHA_MASTER, 2);
  const user = masterData.find(row => String(row[1]).trim() === userEmail);
  return user ? user[0] : null;
}

function logWithTimestamp(message) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMATS.DATETIME);
  console.log(`[${timestamp}] ${message}`);
}

// =================================================================================
// === キャッシュクリア機能 ===
// =================================================================================
function clearScriptCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll();
    SpreadsheetApp.getActiveSpreadsheet().toast('スクリプトのキャッシュをクリアしました。', '完了', 3);
    logWithTimestamp("スクリプトのキャッシュがクリアされました。");
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`キャッシュのクリア中にエラーが発生しました: ${e.message}`, 'エラー', 5);
    Logger.log(`キャッシュのクリア中にエラー: ${e.message}`);
  }
}