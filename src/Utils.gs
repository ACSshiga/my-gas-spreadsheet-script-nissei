/**
 * Utils.gs
 * 汎用ユーティリティ関数を定義します。
 * システム全体で使用される共通の補助関数です。
 */

// =================================================================================
// === 日付関連ユーティリティ ===
// =================================================================================

/**
 * 日付を比較用の "YYYY-MM-DD" 形式の文字列に変換します。
 * @param {Date} date - 変換する日付オブジェクト
 * @return {string|null} - フォーマットされた文字列、または無効な日付の場合はnull
 */
function formatDateForComparison(date) {
  if (!isValidDate(date)) return null;
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

/**
 * 日本の祝日を取得し、Setとして返します。結果はキャッシュされます。
 * @param {number} year - 取得対象の年
 * @return {Set<string>} - 祝日の日付文字列 (YYYY-MM-DD) のセット
 */
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

    cache.put(cacheKey, JSON.stringify([...holidays]), 21600); // 6時間キャッシュ
    return holidays;
  } catch (e) {
    console.error("祝日取得エラー:", e);
    return new Set();
  }
}

/**
 * 指定された日付が休日（土日または祝日）かどうかを判定します。
 */
function isHoliday(date, holidaySet) {
  if (!isValidDate(date)) return false;
  const dayOfWeek = date.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) return true; // 0:日曜, 6:土曜

  return holidaySet.has(formatDateForComparison(date));
}

/**
 * 値が有効な日付オブジェクトであるか検証します。
 */
function isValidDate(value) {
  return value instanceof Date && !isNaN(value.getTime());
}

// =================================================================================
// === 文字列・データ処理ユーティリティ ===
// =================================================================================

/**
 * 値を安全に文字列に変換し、前後の空白を削除します。
 */
function safeTrim(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

/**
 * HYPERLINK関数で安全に使用できるよう、文字列内のダブルクォートをエスケープします。
 */
function createHyperlinkFormula(url, displayText) {
  if (!url) return "";
  const safeUrl = String(url).replace(/"/g, '""');
  const safeText = String(displayText || url).replace(/"/g, '""');
  return `=HYPERLINK("${safeUrl}", "${safeText}")`;
}

/**
 * 値を数値に変換します。変換できない場合は0を返します。
 */
function toNumber(value) {
  const num = parseFloat(value);
  return isNaN(num) ? 0 : num;
}


// =================================================================================
// === マスタデータ取得ユーティリティ ===
// =================================================================================
/**
 * 指定されたマスタシートから値のリストを取得する汎用関数です。
 * @param {string} masterSheetName - マスタシートの名前
 * @param {number} [numColumns=1] - 取得する列数
 * @return {any[][]} - マスタシートから取得した値の2次元配列
 */
function getMasterData(masterSheetName, numColumns = 1) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `master_${masterSheetName}_${numColumns}`;
  const cached = cache.get(cacheKey);

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) { /* ignore */ }
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterSheetName);
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, numColumns).getValues();
  const filteredValues = values.filter(row => row[0] !== ""); // 1列目が空の行は除外
  
  cache.put(cacheKey, JSON.stringify(filteredValues), 3600); // 1時間キャッシュ
  return filteredValues;
}

/**
* メールアドレスから担当者マスタを検索し、対応する担当者名を返します。
* ★メールアドレスの前後の空白を無視するように修正
*/
function getTantoushaNameByEmail(email) {
  if (!email) return null;
  const userEmail = email.trim(); // ★自分のメールアドレスの空白を削除
  const masterData = getMasterData(CONFIG.SHEETS.TANTOUSHA_MASTER, 2);
  const user = masterData.find(row => String(row[1]).trim() === userEmail); // ★マスタ側の空白も削除して比較
  return user ? user[0] : null;
}

/**
 * タイムスタンプ付きでLoggerにログを出力します。
 */
function logWithTimestamp(message) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMATS.DATETIME);
  console.log(`[${timestamp}] ${message}`);
}