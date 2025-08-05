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
*/
function getTantoushaNameByEmail(email) {
  if (!email) return null;
  const masterData = getMasterData(CONFIG.SHEETS.TANTOUSHA_MASTER, 2);
  const user = masterData.find(row => row[1] === email); // 2列目(B列)がメールアドレス
  return user ? user[0] : null; // 1列目(A列)の担当者名を返す
}

/**
 * タイムスタンプ付きでLoggerにログを出力します。
 */
function logWithTimestamp(message) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMATS.DATETIME);
  console.log(`[${timestamp}] ${message}`);
}/**
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
    return JSON.parse(cached);
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
*/
function getTantoushaNameByEmail(email) {
  const masterData = getMasterData(CONFIG.SHEETS.TANTOUSHA_MASTER, 2);
  const user = masterData.find(row => row[1] === email); // 2列目(B列)がメールアドレス
  return user ? user[0] : null; // 1列目(A列)の担当者名を返す
}

/**
 * タイムスタンプ付きでLoggerにログを出力します。
 */
function logWithTimestamp(message) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMATS.DATETIME);
  console.log(`[${timestamp}] ${message}`);
}/**
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
 * @param {Date} date - 判定する日付
 * @param {Set<string>} holidaySet - 事前に取得した祝日のセット
 * @return {boolean} - 休日の場合はtrue
 */
function isHoliday(date, holidaySet) {
  if (!isValidDate(date)) return false;
  const dayOfWeek = date.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) return true; // 0:日曜, 6:土曜

  return holidaySet.has(formatDateForComparison(date));
}

/**
 * 値が有効な日付オブジェクトであるか検証します。
 * @param {*} value - 検証する値
 * @return {boolean} - 有効な日付の場合はtrue
 */
function isValidDate(value) {
  return value instanceof Date && !isNaN(value.getTime());
}

// =================================================================================
// === 文字列・データ処理ユーティリティ ===
// =================================================================================

/**
 * 値を安全に文字列に変換し、前後の空白を削除します。
 * @param {*} value - 変換する値
 * @return {string} - トリムされた文字列
 */
function safeTrim(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

/**
 * HYPERLINK関数で安全に使用できるよう、文字列内のダブルクォートをエスケープします。
 * @param {string} str - エスケープする文字列
 * @return {string} - エスケープされた文字列
 */
function escapeForHyperlink(str) {
  return String(str).replace(/"/g, '""');
}

/**
 * HYPERLINK関数の文字列を生成します。
 * @param {string} url - リンク先のURL
 * @param {string} displayText - 表示するテキスト
 * @return {string} - HYPERLINK関数の文字列
 */
function createHyperlinkFormula(url, displayText) {
  if (!url) return "";
  const safeUrl = escapeForHyperlink(url);
  const safeText = escapeForHyperlink(displayText || url);
  return `=HYPERLINK("${safeUrl}", "${safeText}")`;
}

/**
 * 値を数値に変換します。変換できない場合は0を返します。
 * @param {*} value - 変換する値
 * @return {number} - 変換後の数値
 */
function toNumber(value) {
  const num = Number(value);
  return isNaN(num) ? 0 : num;
}


// =================================================================================
// === マスタデータ取得ユーティリティ ===
// =================================================================================
/**
 * 指定されたマスタシートから値のリストを取得する汎用関数です。
 * @param {string} masterSheetName - マスタシートの名前
 * @return {string[]} - マスタシートの1列目から取得した値の配列
 */
function getMasterData(masterSheetName) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `master_${masterSheetName}`;
  const cached = cache.get(cacheKey);

  if (cached) {
    return JSON.parse(cached);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterSheetName);
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String);
  cache.put(cacheKey, JSON.stringify(values), 3600); // 1時間キャッシュ
  return values;
}


/**
 * タイムスタンプ付きでLoggerにログを出力します。
 * @param {string} message - ログメッセージ
 */
function logWithTimestamp(message) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMATS.DATETIME);
  console.log(`[${timestamp}] ${message}`);
}/**
 * Utils.gs
 * 汎用ユーティリティ関数
 * システム全体で使用される共通関数を定義
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
 * @param {Date} startDate - 取得開始日
 * @param {Date} endDate - 取得終了日
 * @return {Set<string>} - 祝日の日付文字列 (YYYY-MM-DD) のセット
 */
function getJapaneseHolidays(startDate, endDate) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `holidays_${formatDateForComparison(startDate)}_${formatDateForComparison(endDate)}`;
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

    const events = calendar.getEvents(startDate, endDate);
    const holidays = events.map(e => formatDateForComparison(e.getStartTime())).filter(d => d);
    const holidaySet = new Set(holidays);

    cache.put(cacheKey, JSON.stringify([...holidaySet]), BATCH_CONFIG.CACHE_EXPIRATION * 10); // 祝日は長めにキャッシュ
    return holidaySet;
  } catch (e) {
    console.error("祝日取得エラー:", e);
    return new Set();
  }
}

/**
 * 指定された日付が休日（土日または祝日）かどうかを判定します。
 * @param {Date} date - 判定する日付
 * @param {Set<string>} holidaySet - 事前に取得した祝日のセット
 * @return {boolean} - 休日の場合はtrue
 */
function isHoliday(date, holidaySet) {
  if (!isValidDate(date)) return false;

  const dayOfWeek = date.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) return true; // 0:日曜, 6:土曜

  if (holidaySet && holidaySet.has(formatDateForComparison(date))) return true;

  return false;
}

// =================================================================================
// === 文字列処理ユーティリティ ===
// =================================================================================

/**
 * 値を安全に文字列に変換し、前後の空白を削除します。
 * @param {*} value - 変換する値
 * @return {string} - トリムされた文字列
 */
function safeTrim(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

/**
 * HYPERLINK関数で安全に使用できるよう、文字列内のダブルクォートをエスケープします。
 * @param {string} str - エスケープする文字列
 * @return {string} - エスケープされた文字列
 */
function escapeForHyperlink(str) {
  return String(str).replace(/"/g, '""');
}

/**
 * HYPERLINK関数の文字列を生成します。
 * @param {string} url - リンク先のURL
 * @param {string} displayText - 表示するテキスト
 * @return {string} - HYPERLINK関数の文字列
 */
function createHyperlinkFormula(url, displayText) {
  if (!url) return "";
  const safeUrl = escapeForHyperlink(url);
  const safeText = escapeForHyperlink(displayText || url);
  return `=HYPERLINK("${safeUrl}", "${safeText}")`;
}

/**
 * 機種名からシリーズ名を抽出します (例: "NEX140V-18E" -> "NEX140")。
 * @param {string} modelString - 機種文字列
 * @return {string|null} - 抽出されたシリーズ名、またはnull
 */
function extractSeriesPlusInitialNumber(modelString) {
  if (!modelString || typeof modelString !== 'string') return null;

  const relevantModelString = modelString.toUpperCase().trim().replace(/^[^A-Z]+/, "");
  if (!relevantModelString) return null;

  const match = relevantModelString.match(/^([A-Z]{2,})(\d+)/);
  if (match) {
    return match[1] + match[2];
  }
  return null;
}

// =================================================================================
// === 数値・データ検証ユーティリティ ===
// =================================================================================

/**
 * 値を数値に変換します。変換できない場合は0を返します。
 * @param {*} value - 変換する値
 * @return {number} - 変換後の数値
 */
function toNumber(value) {
  const num = Number(value);
  return isNaN(num) ? 0 : num;
}

/**
 * 値が有効な日付オブジェクトであるか検証します。
 * @param {*} value - 検証する値
 * @return {boolean} - 有効な日付の場合はtrue
 */
function isValidDate(value) {
  return value instanceof Date && !isNaN(value.getTime());
}

/**
 * タイムスタンプ付きでLoggerにログを出力します。
 * @param {string} message - ログメッセージ
 */
function logWithTimestamp(message) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMATS.DATETIME);
  console.log(`[${timestamp}] ${message}`);
}/**
 * Utils.gs
 * 汎用ユーティリティ関数
 * システム全体で使用される共通関数を定義
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
 * @param {Date} startDate - 取得開始日
 * @param {Date} endDate - 取得終了日
 * @return {Set<string>} - 祝日の日付文字列 (YYYY-MM-DD) のセット
 */
function getJapaneseHolidays(startDate, endDate) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `holidays_${formatDateForComparison(startDate)}_${formatDateForComparison(endDate)}`;
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
    
    const events = calendar.getEvents(startDate, endDate);
    const holidays = events.map(e => formatDateForComparison(e.getStartTime())).filter(d => d);
    const holidaySet = new Set(holidays);
    
    cache.put(cacheKey, JSON.stringify([...holidaySet]), BATCH_CONFIG.CACHE_EXPIRATION * 10); // 祝日は長めにキャッシュ
    return holidaySet;
  } catch (e) {
    console.error("祝日取得エラー:", e);
    return new Set();
  }
}

/**
 * 指定された日付が休日（土日または祝日）かどうかを判定します。
 * @param {Date} date - 判定する日付
 * @param {Set<string>} holidaySet - 事前に取得した祝日のセット
 * @return {boolean} - 休日の場合はtrue
 */
function isHoliday(date, holidaySet) {
  if (!isValidDate(date)) return false;
  
  const dayOfWeek = date.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) return true; // 0:日曜, 6:土曜
  
  if (holidaySet && holidaySet.has(formatDateForComparison(date))) return true;
  
  return false;
}

// =================================================================================
// === 文字列処理ユーティリティ ===
// =================================================================================

/**
 * 値を安全に文字列に変換し、前後の空白を削除します。
 * @param {*} value - 変換する値
 * @return {string} - トリムされた文字列
 */
function safeTrim(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

/**
 * HYPERLINK関数で安全に使用できるよう、文字列内のダブルクォートをエスケープします。
 * @param {string} str - エスケープする文字列
 * @return {string} - エスケープされた文字列
 */
function escapeForHyperlink(str) {
  return String(str).replace(/"/g, '""');
}

/**
 * HYPERLINK関数の文字列を生成します。
 * @param {string} url - リンク先のURL
 * @param {string} displayText - 表示するテキスト
 * @return {string} - HYPERLINK関数の文字列
 */
function createHyperlinkFormula(url, displayText) {
  if (!url) return "";
  const safeUrl = escapeForHyperlink(url);
  const safeText = escapeForHyperlink(displayText || url);
  return `=HYPERLINK("${safeUrl}", "${safeText}")`;
}

/**
 * 機種名からシリーズ名を抽出します (例: "NEX140V-18E" -> "NEX140")。
 * @param {string} modelString - 機種文字列
 * @return {string|null} - 抽出されたシリーズ名、またはnull
 */
function extractSeriesPlusInitialNumber(modelString) {
  if (!modelString || typeof modelString !== 'string') return null;
  
  const relevantModelString = modelString.toUpperCase().trim().replace(/^[^A-Z]+/, "");
  if (!relevantModelString) return null;
  
  const match = relevantModelString.match(/^([A-Z]{2,})(\d+)/);
  if (match) {
    return match[1] + match[2];
  }
  return null;
}

// =================================================================================
// === 数値・データ検証ユーティリティ ===
// =================================================================================

/**
 * 値を数値に変換します。変換できない場合は0を返します。
 * @param {*} value - 変換する値
 * @return {number} - 変換後の数値
 */
function toNumber(value) {
  const num = Number(value);
  return isNaN(num) ? 0 : num;
}

/**
 * 値が有効な日付オブジェクトであるか検証します。
 * @param {*} value - 検証する値
 * @return {boolean} - 有効な日付の場合はtrue
 */
function isValidDate(value) {
  return value instanceof Date && !isNaN(value.getTime());
}

/**
 * タイムスタンプ付きでLoggerにログを出力します。
 * @param {string} message - ログメッセージ
 */
function logWithTimestamp(message) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMATS.DATETIME);
  console.log(`[${timestamp}] ${message}`);
}/**
 * Utils.gs
 * 汎用ユーティリティ関数
 * システム全体で使用される共通関数を定義
 */

// =================================================================================
// === 日付関連ユーティリティ ===
// =================================================================================

/**
 * 日付を比較用の文字列形式に変換
 * @param {Date} date - 変換する日付
 * @return {string|null} YYYY-MM-DD形式の文字列、無効な日付の場合はnull
 */
function formatDateForComparison(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return null;
  }
  const year = date.getFullYear();
  const month = ('0' + (date.getMonth() + 1)).slice(-2);
  const day = ('0' + date.getDate()).slice(-2);
  return `${year}-${month}-${day}`;
}

/**
 * 日付をフォーマット
 * @param {Date} date - フォーマットする日付
 * @param {string} format - フォーマット文字列（DATE_FORMATSから選択）
 * @return {string} フォーマットされた日付文字列
 */
function formatDate(date, format) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return "";
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), format);
}

/**
 * 日本の祝日を取得
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Set<string>} 祝日のセット（YYYY-MM-DD形式）
 */
function getJapaneseHolidays(startDate, endDate) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `holidays_${formatDateForComparison(startDate)}_${formatDateForComparison(endDate)}`;
  const cached = cache.get(cacheKey);
  
  if (cached) {
    return new Set(JSON.parse(cached));
  }
  
  try {
    const calendar = CalendarApp.getCalendarById(CONFIG.HOLIDAY_CALENDAR_ID);
    if (!calendar) {
      console.warn("祝日カレンダーが見つかりません");
      return new Set();
    }
    
    const events = calendar.getEvents(startDate, endDate);
    const holidays = events.map(e => formatDateForComparison(e.getStartTime())).filter(d => d);
    const holidaySet = new Set(holidays);
    
    cache.put(cacheKey, JSON.stringify([...holidaySet]), BATCH_CONFIG.CACHE_EXPIRATION);
    return holidaySet;
  } catch (e) {
    console.error("祝日取得エラー:", e);
    return new Set();
  }
}

/**
 * 日付が休日（土日祝）かどうかを判定
 * @param {Date} date - 判定する日付
 * @param {Set<string>} holidaySet - 祝日のセット（オプション）
 * @return {boolean} 休日の場合true
 */
function isHoliday(date, holidaySet = null) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return false;
  }
  
  const dayOfWeek = date.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    return true; // 土日
  }
  
  if (holidaySet) {
    return holidaySet.has(formatDateForComparison(date));
  }
  
  return false;
}

// =================================================================================
// === 文字列処理ユーティリティ ===
// =================================================================================

/**
 * 文字列を安全にトリミング
 * @param {*} value - トリミングする値
 * @return {string} トリミングされた文字列
 */
function safeTrim(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).trim();
}

/**
 * ハイパーリンク用に文字列内のダブルクォートをエスケープ
 * @param {string} str - エスケープする文字列
 * @return {string} エスケープされた文字列
 */
function escapeForHyperlink(str) {
  return String(str).replace(/"/g, '""');
}

/**
 * HYPERLINK関数の文字列を生成
 * @param {string} url - リンクURL
 * @param {string} displayText - 表示テキスト
 * @return {string} HYPERLINK関数の文字列
 */
function createHyperlinkFormula(url, displayText) {
  if (!url) return "";
  const safeUrl = escapeForHyperlink(url);
  const safeText = escapeForHyperlink(displayText || url);
  return `=HYPERLINK("${safeUrl}", "${safeText}")`;
}

/**
 * 機種名からシリーズ名を抽出
 * @param {string} modelString - 機種文字列
 * @param {string} kibanString - 機番文字列（オプション）
 * @return {string|null} 抽出されたシリーズ名
 */
function extractSeriesPlusInitialNumber(modelString, kibanString = "") {
  if (!modelString || typeof modelString !== 'string') {
    return null;
  }
  
  let relevantModelString = modelString.toUpperCase().trim();
  const kibanUpper = kibanString ? String(kibanString).toUpperCase().trim() : "";
  
  if (kibanUpper && relevantModelString.startsWith(kibanUpper)) {
    relevantModelString = relevantModelString.substring(kibanUpper.length).replace(/^[^A-ZⅣⅤⅢⅡⅠV]+/, "");
  } else {
    relevantModelString = relevantModelString.replace(/^[^A-ZⅣⅤⅢⅡⅠV]+/, "");
  }
  
  if (!relevantModelString) return null;
  
  const match = relevantModelString.match(/^([A-Z]{2,})(\d+)/i);
  if (match) {
    const extractedName = match[1] + match[2];
    if (relevantModelString.toUpperCase().startsWith(extractedName.toUpperCase()) && 
        match[1].length >= 2 && match[2].length >= 1) {
      return extractedName;
    }
  }
  
  return null;
}

// =================================================================================
// === 数値処理ユーティリティ ===
// =================================================================================

/**
 * 値を数値に変換（デフォルト値付き）
 * @param {*} value - 変換する値
 * @param {number} defaultValue - デフォルト値（デフォルト: 0）
 * @return {number} 数値
 */
function toNumber(value, defaultValue = 0) {
  const num = Number(value);
  return isNaN(num) ? defaultValue : num;
}

/**
 * 数値を指定した小数点以下桁数に丸める
 * @param {number} value - 丸める値
 * @param {number} decimals - 小数点以下桁数
 * @return {number} 丸められた値
 */
function roundTo(value, decimals = 2) {
  const factor = Math.pow(10, decimals);
  return Math.round(value * factor) / factor;
}

// =================================================================================
// === 配列・オブジェクト処理ユーティリティ ===
// =================================================================================

/**
 * 2次元配列を1次元配列に変換（特定の列のみ）
 * @param {Array<Array>} array2D - 2次元配列
 * @param {number} columnIndex - 抽出する列のインデックス（0-based）
 * @return {Array} 1次元配列
 */
function extractColumn(array2D, columnIndex) {
  return array2D.map(row => row[columnIndex]);
}

/**
 * 配列から重複を除去
 * @param {Array} array - 処理する配列
 * @return {Array} 重複が除去された配列
 */
function uniqueArray(array) {
  return [...new Set(array)];
}

/**
 * オブジェクトの深いコピーを作成
 * @param {Object} obj - コピーするオブジェクト
 * @return {Object} コピーされたオブジェクト
 */
function deepCopy(obj) {
  return JSON.parse(JSON.stringify(obj));
}

/**
 * 2つの配列の差分を取得
 * @param {Array} array1 - 配列1
 * @param {Array} array2 - 配列2
 * @return {Array} array1にあってarray2にない要素
 */
function arrayDifference(array1, array2) {
  const set2 = new Set(array2);
  return array1.filter(x => !set2.has(x));
}

// =================================================================================
// === バリデーションユーティリティ ===
// =================================================================================

/**
 * 値が空かどうかをチェック
 * @param {*} value - チェックする値
 * @return {boolean} 空の場合true
 */
function isEmpty(value) {
  return value === null || 
         value === undefined || 
         value === "" || 
         (Array.isArray(value) && value.length === 0) ||
         (typeof value === 'object' && Object.keys(value).length === 0);
}

/**
 * 有効な日付かどうかをチェック
 * @param {*} value - チェックする値
 * @return {boolean} 有効な日付の場合true
 */
function isValidDate(value) {
  return value instanceof Date && !isNaN(value.getTime());
}

/**
 * 有効なメールアドレスかどうかをチェック
 * @param {string} email - チェックするメールアドレス
 * @return {boolean} 有効なメールアドレスの場合true
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

// =================================================================================
// === ログ・デバッグユーティリティ ===
// =================================================================================

/**
 * タイムスタンプ付きでログを出力
 * @param {string} message - ログメッセージ
 * @param {string} level - ログレベル（INFO, WARN, ERROR）
 */
function logWithTimestamp(message, level = "INFO") {
  const timestamp = formatDate(new Date(), DATE_FORMATS.DATETIME);
  console.log(`[${timestamp}] [${level}] ${message}`);
}

/**
 * 処理時間を計測するタイマークラス
 */
class Timer {
  constructor(name = "Process") {
    this.name = name;
    this.startTime = new Date().getTime();
  }
  
  /**
   * 経過時間をログに出力
   * @param {string} message - 追加メッセージ
   */
  log(message = "") {
    const elapsed = new Date().getTime() - this.startTime;
    const suffix = message ? ` - ${message}` : "";
    logWithTimestamp(`${this.name}: ${elapsed}ms${suffix}`, "INFO");
  }
  
  /**
   * タイマーをリセット
   */
  reset() {
    this.startTime = new Date().getTime();
  }
}

// =================================================================================
// === エラーハンドリングユーティリティ ===
// =================================================================================

/**
 * エラーセーフな関数実行
 * @param {Function} func - 実行する関数
 * @param {*} defaultValue - エラー時のデフォルト値
 * @param {string} errorMessage - エラーメッセージ
 * @return {*} 関数の戻り値またはデフォルト値
 */
function tryCatch(func, defaultValue = null, errorMessage = "エラーが発生しました") {
  try {
    return func();
  } catch (error) {
    logWithTimestamp(`${errorMessage}: ${error.message}`, "ERROR");
    return defaultValue;
  }
}

/**
 * リトライ機能付き関数実行
 * @param {Function} func - 実行する関数
 * @param {number} maxRetries - 最大リトライ回数
 * @param {number} delay - リトライ間隔（ミリ秒）
 * @return {*} 関数の戻り値
 */
function retryOperation(func, maxRetries = 3, delay = 1000) {
  let lastError;
  
  for (let i = 0; i <= maxRetries; i++) {
    try {
      return func();
    } catch (error) {
      lastError = error;
      if (i < maxRetries) {
        logWithTimestamp(`リトライ ${i + 1}/${maxRetries}: ${error.message}`, "WARN");
        Utilities.sleep(delay);
      }
    }
  }
  
  throw lastError;
}