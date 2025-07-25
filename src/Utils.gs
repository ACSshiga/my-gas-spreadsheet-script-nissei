/**
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