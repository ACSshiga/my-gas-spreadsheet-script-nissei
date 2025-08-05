/**
 * Config.gs
 * システム全体の設定定数を管理します。
 * 将来の仕様変更に対応できるよう、設定はここに集約されています。
 */

// =================================================================================
// === グローバル設定 ===
// =================================================================================
const CONFIG = {
  // Google DriveフォルダID
  FOLDERS: {
    REFERENCE_MATERIAL_PARENT: "124OR71hkr2jeT-5esv0GHAeZn83fAvYc", // 製番資料の親フォルダ
    SERIES_MODEL_PARENT:       "1XdiYBWiixF_zOSScT7UKUhCQkye3MLNJ", // STD資料の親フォルダ
    BACKUP_PARENT:             "1HCDyjF_Kw2jlzN491uZK1X6QeXFuOWSl"  // バックアップの親フォルダ
  },

  // シート名
  SHEETS: {
    MAIN:             "メインシート",
    INPUT_PREFIX:     "工数_",
    TANTOUSHA_MASTER: "担当者マスタ",
    SAGYOU_KUBUN_MASTER: "作業区分マスタ",
    SHINCHOKU_MASTER: "進捗マスタ",
    TOIAWASE_MASTER:  "問い合わせマスタ"
  },

  // 色設定
  COLORS: {
    DEFAULT_BACKGROUND: '#ffffff', // デフォルトの背景色
    WEEKEND_HOLIDAY:    '#f2f2f2', // 休日の背景色
    FILTER_HIGHLIGHT:   '#fef7e0'  // フィルタリング中のヘッダー色
  },

  // データ開始行
  DATA_START_ROW: {
    MAIN: 2,  // メインシートのデータは2行目から
    INPUT: 3  // 工数シートのデータは3行目から
  },

  // バックアップ設定
  BACKUP: {
    KEEP_COUNT: 5,               // 保持するバックアップ数
    PREFIX:     "【Backup】"     // バックアップファイル名の接頭辞
  },

  // 日本の祝日カレンダーID
  HOLIDAY_CALENDAR_ID: 'ja.japanese#holiday@group.v.calendar.google.com',
};

// =================================================================================
// === 動的な色設定（マスタ連動） ===
// =================================================================================
// マスタに新しいステータスを追加した場合、ここに色定義を追加するだけで対応できます。
const PROGRESS_COLORS = new Map([
  ["未着手", "#ffcdd2"],
  ["仕掛中", "#e1bee7"],
  ["保留",   "#c5cae9"],
  ["完了",   "#c8e6c9"],
]);

const TANTOUSHA_COLORS = new Map([
  ["志賀", "#ffccbc"],
  ["遠藤", "#dcedc8"],
  ["小板橋", "#b2ebf2"],
]);

const TOIAWASE_COLORS = new Map([
  ["問合済", "#ffecb3"],
  ["回答済", "#d1c4e9"],
]);


// =================================================================================
// === シート列定義（ヘッダー名ベース） ===
// =================================================================================
/**
 * メインシートの列定義
 */
const MAIN_SHEET_HEADERS = {
  MGMT_NO:          "管理No",
  SAGYOU_KUBUN:     "作業区分",
  KIBAN:            "機番",
  MODEL:            "機種",
  KIBAN_URL:        "機番(リンク)",
  SERIES_URL:       "STD資料(リンク)",
  REFERENCE_KIBAN:  "参考製番",
  CIRCUIT_DIAGRAM_NO: "回路図番",
  TOIAWASE:         "問い合わせ",
  TEMP_CODE:        "仮コード",
  TANTOUSHA:        "担当者",
  DESTINATION:      "納入先",
  PLANNED_HOURS:    "予定工数",
  ACTUAL_HOURS:     "実績工数",
  PROGRESS:         "進捗",
  START_DATE:       "仕掛日",
  COMPLETE_DATE:    "完了日",
  DRAWING_DEADLINE: "作図期限",
  PROGRESS_EDITOR:  "進捗記入者",
  UPDATE_TS:        "更新日時",
  OVERRUN_REASON:   "係り超過理由",
  NOTES:            "注意点",
  REMARKS:          "備考"
};

/**
 * 工数シートの列定義
 */
const INPUT_SHEET_HEADERS = {
  MGMT_NO:      "管理No.",
  SAGYOU_KUBUN: "作業区分",
  KIBAN:        "機番",
  PROGRESS:     "進捗",
  PLANNED_HOURS:"予定工数",
  ACTUAL_HOURS_SUM: "実績工数合計",
  // G列以降は動的な日付列
};


// =================================================================================
// === ユーティリティ関数（Config内） ===
// =================================================================================
/**
 * シートのヘッダー行を読み取り、各列が何番目にあるかを動的に取得します。
 */
function getColumnIndices(sheet, headerDef) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `col_indices_${sheet.getSheetId()}`;
  const cached = cache.get(cacheKey);

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e) {
      // キャッシュが無効な場合は再取得
    }
  }

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const indices = {};
  for (const [key, headerName] of Object.entries(headerDef)) {
    const index = headerRow.indexOf(headerName) + 1;
    if (index > 0) {
      indices[key] = index;
    }
  }

  cache.put(cacheKey, JSON.stringify(indices), 21600); // 6時間キャッシュ
  return indices;
}

/**
 * 色定義マップから、指定されたキーに対応する色コードを取得します。
 */
function getColor(colorMap, key) {
  return colorMap.get(key) || CONFIG.COLORS.DEFAULT_BACKGROUND;
}

// =================================================================================
// === 日付フォーマット定義 ===
// =================================================================================
const DATE_FORMATS = {
  DATE_ONLY:        "yyyy/MM/dd",
  DATETIME:         "yyyy-MM-dd HH:mm:ss",
  MONTH_DAY:        "M/d",
  BACKUP_TIMESTAMP: "yyyy-MM-dd_HH-mm"
};/**
 * Config.gs
 * システム全体の設定定数を管理します。
 * 将来の仕様変更に対応できるよう、設定はここに集約されています。
 */

// =================================================================================
// === グローバル設定 ===
// =================================================================================
const CONFIG = {
  // Google DriveフォルダID
  FOLDERS: {
    REFERENCE_MATERIAL_PARENT: "124OR71hkr2jeT-5esv0GHAeZn83fAvYc", // 製番資料の親フォルダ
    SERIES_MODEL_PARENT:       "1XdiYBWiixF_zOSScT7UKUhCQkye3MLNJ", // STD資料の親フォルダ
    BACKUP_PARENT:             "1HCDyjF_Kw2jlzN491uZK1X6QeXFuOWSl"  // バックアップの親フォルダ
  },

  // シート名
  SHEETS: {
    MAIN:             "メインシート",
    INPUT_PREFIX:     "工数_",
    TANTOUSHA_MASTER: "担当者マスタ",
    SAGYOU_KUBUN_MASTER: "作業区分マスタ",
    SHINCHOKU_MASTER: "進捗マスタ",
    TOIAWASE_MASTER:  "問い合わせマスタ"
  },

  // 色設定
  COLORS: {
    DEFAULT_BACKGROUND: '#ffffff', // デフォルトの背景色
    WEEKEND_HOLIDAY:    '#e8eaed', // 休日の背景色
    FILTER_HIGHLIGHT:   '#fef7e0'  // フィルタリング中のヘッダー色
  },

  // バックアップ設定
  BACKUP: {
    KEEP_COUNT: 5,               // 保持するバックアップ数
    PREFIX:     "【Backup】"     // バックアップファイル名の接頭辞
  },

  // 日本の祝日カレンダーID
  HOLIDAY_CALENDAR_ID: 'ja.japanese#holiday@group.v.calendar.google.com',
};

// =================================================================================
// === 動的な色設定（マスタ連動） ===
// =================================================================================
// マスタに新しいステータスを追加した場合、ここに色定義を追加するだけで対応できます。
const PROGRESS_COLORS = new Map([
  ["未着手", "#ffcdd2"],
  ["仕掛中", "#e1bee7"],
  ["保留",   "#c5cae9"],
  ["完了",   "#c8e6c9"],
  ["レビュー中", "#b3e5fc"], // 将来の追加を見越したサンプル
]);

const TANTOUSHA_COLORS = new Map([
  ["志賀", "#ffccbc"],
  ["遠藤", "#dcedc8"],
  ["小板橋", "#b2ebf2"],
]);

const TOIAWASE_COLORS = new Map([
  ["問合済", "#ffecb3"],
  ["回答済", "#d1c4e9"],
]);


// =================================================================================
// === シート列定義（ヘッダー名ベース） ===
// =================================================================================
/**
 * メインシートの列定義
 * この定義を変更するだけで、スクリプトは列の追加・順序変更に自動で追従します。
 */
const MAIN_SHEET_HEADERS = {
  MGMT_NO:          "管理No",
  SAGYOU_KUBUN:     "作業区分",
  KIBAN:            "機番",
  MODEL:            "機種",
  KIBAN_URL:        "機番(リンク)",
  SERIES_URL:       "STD資料(リンク)",
  REFERENCE_KIBAN:  "参考製番",
  CIRCUIT_DIAGRAM_NO: "回路図番",
  TOIAWASE:         "問い合わせ",
  TEMP_CODE:        "仮コード",
  TANTOUSHA:        "担当者",
  DESTINATION:      "納入先",
  PLANNED_HOURS:    "予定工数",
  ACTUAL_HOURS:     "実績工数",
  PROGRESS:         "進捗",
  START_DATE:       "仕掛日",
  COMPLETE_DATE:    "完了日",
  DRAWING_DEADLINE: "作図期限",
  PROGRESS_EDITOR:  "進捗記入者",
  UPDATE_TS:        "更新日時",
  OVERRUN_REASON:   "係り超過理由",
  NOTES:            "注意点",
  REMARKS:          "備考"
};

/**
 * 工数シートの列定義
 * 入力に特化した、シンプルな構成です。
 */
const INPUT_SHEET_HEADERS = {
  MGMT_NO:      "管理No.",
  SAGYOU_KUBUN: "作業区分",
  KIBAN:        "機番",
  PROGRESS:     "進捗",
  // E列以降は動的な日付列
};


// =================================================================================
// === ユーティリティ関数（Config内） ===
// =================================================================================
/**
 * シートのヘッダー行を読み取り、各列が何番目にあるかを動的に取得します。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {Object} headerDef - ヘッダー定義オブジェクト
 * @return {Object} ヘッダー名をキー、列インデックス(1-based)を値とするオブジェクト
 */
function getColumnIndices(sheet, headerDef) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `col_indices_${sheet.getSheetId()}`;
  const cached = cache.get(cacheKey);

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e) {
      // キャッシュが無効な場合は再取得
    }
  }

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const indices = {};
  for (const [key, headerName] of Object.entries(headerDef)) {
    const index = headerRow.indexOf(headerName) + 1;
    if (index > 0) {
      indices[key] = index;
    }
  }

  cache.put(cacheKey, JSON.stringify(indices), 21600); // 6時間キャッシュ
  return indices;
}

/**
 * 色定義マップから、指定されたキーに対応する色コードを取得します。
 * @param {Map<string, string>} colorMap - 色の対応表
 * @param {string} key - 色を取得したいキー（例: "未着手"）
 * @return {string} 色コード（見つからない場合はデフォルト色）
 */
function getColor(colorMap, key) {
  return colorMap.get(key) || CONFIG.COLORS.DEFAULT_BACKGROUND;
}

// =================================================================================
// === 日付フォーマット定義 ===
// =================================================================================
const DATE_FORMATS = {
  DATE_ONLY:        "yyyy/MM/dd",
  DATETIME:         "yyyy-MM-dd HH:mm:ss",
  MONTH_DAY:        "M/d",
  BACKUP_TIMESTAMP: "yyyy-MM-dd_HH-mm"
};/**
 * Config.gs
 * システム全体の設定定数を管理
 */

// =================================================================================
// === グローバル設定 ===
// =================================================================================
const CONFIG = {
  // Google DriveフォルダID
  FOLDERS: {
    REFERENCE_MATERIAL_PARENT: "124OR71hkr2jeT-5esv0GHAeZn83fAvYc",
    SERIES_MODEL_PARENT: "1XdiYBWiixF_zOSScT7UKUhCQkye3MLNJ",
    BACKUP_PARENT: "1HCDyjF_Kw2jlzN491uZK1X6QeXFuOWSl"
  },

  // シート名
  SHEETS: {
    MAIN: "メインシート",
    INPUT_PREFIX: "工数_",
    TANTOUSHA_MASTER: "担当者マスタ", // ★担当者マスタシートを定義
  },

  // 色設定
  COLORS: {
    DUPLICATE_HIGHLIGHT: '#e8eaed',
    DEFAULT_BACKGROUND: '#ffffff',
    WEEKEND_HOLIDAY: '#e8eaed'
  },

  // バックアップ設定
  BACKUP: {
    KEEP_COUNT: 5,
    PREFIX: "【Backup】"
  },

  // 日本の祝日カレンダーID
  HOLIDAY_CALENDAR_ID: 'ja.japanese#holiday@group.v.calendar.google.com',

  // デフォルト値
  DEFAULTS: {
    PROGRESS: "未着手",
    DUPLICATE_TEXT: "機番重複"
  }
};

// =================================================================================
// === 動的な色設定 ===
// =================================================================================
const PROGRESS_COLORS = new Map([
  ["未着手", "#ffcfc9"],
  ["仕掛中", "#e6cff2"],
  ["保留", "#c6dbe1"],
  ["完了", "#d4edbc"],
  ["機番重複", "#e8eaed"]
]);

const TANTOUSHA_COLORS = new Map([
  ["志賀", "#ffcfc9"],
  ["遠藤", "#d4edbc"],
  ["小板橋", "#bfe1f6"]
]);

const TOIAWASE_COLORS = new Map([
  ["問合済", "#ffcfc9"],
  ["回答済", "#bfe1f6"]
]);

// =================================================================================
// === シート列定義（ヘッダー名ベース） ===
// =================================================================================
/**
 * ★★★ 要望を反映した最終的な列構成 ★★★
 */
const MAIN_SHEET_HEADERS = {
  MGMT_NO: "管理No",
  KIBAN: "機番",
  MODEL: "機種",
  KIBAN_URL: "機番(リンク)",
  SERIES_URL: "STD資料(リンク)",
  REFERENCE_KIBAN: "参考製番",
  CIRCUIT_DIAGRAM_NO: "回路図番",
  TOIAWASE: "問い合わせ",
  TEMP_CODE: "仮コード",
  TANTOUSHA: "担当者",
  DESTINATION: "納入先",
  PLANNED_HOURS_PANEL: "予定工数(盤配)",
  ACTUAL_HOURS_PANEL: "実績工数(盤配)",
  PLANNED_HOURS_WIRE: "予定工数(線加工)",
  ACTUAL_HOURS_WIRE: "実績工数(線加工)",
  PROGRESS_PANEL: "進捗(盤配)",
  PROGRESS_WIRE: "進捗(線加工)",
  START_DATE: "仕掛日",
  COMPLETE_DATE: "完了日",
  DRAWING_DEADLINE: "作図期限",
  PROGRESS_EDITOR: "進捗記入者",
  UPDATE_TS: "更新日時",
  OVERRUN_REASON: "係り超過理由",
  NOTES: "注意点",
  REMARKS: "備考"
};

/**
 * 工数シートの列定義（メインシートと完全に一致させる）
 */
const INPUT_SHEET_HEADERS = MAIN_SHEET_HEADERS;

// =================================================================================
// === ユーティリティ関数（Config内） ===
// =================================================================================
/**
 * ヘッダー行から列インデックスを動的に取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {Object} headerDef - ヘッダー定義オブジェクト
 * @return {Object} ヘッダー名をキー、列インデックスを値とするオブジェクト
 */
function getColumnIndices(sheet, headerDef) {
  const cache = CacheService.getScriptCache();
  // シートIDとシート名でキャッシュキーを生成し、名前変更に対応
  const cacheKey = `col_indices_${sheet.getSheetId()}_${sheet.getName()}`;
  const cached = cache.get(cacheKey);

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e) { /* パース失敗時は再生成 */ }
  }

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const indices = {};
  for (const [key, headerName] of Object.entries(headerDef)) {
    const index = headerRow.indexOf(headerName) + 1;
    if (index > 0) {
      indices[key] = index;
    } else {
      Logger.log(`警告: ヘッダー "${headerName}" がシート "${sheet.getName()}" に見つかりません。`);
    }
  }

  cache.put(cacheKey, JSON.stringify(indices), 21600); // 6時間キャッシュ
  return indices;
}

/**
 * 色を取得する汎用関数
 */
function getColor(colorMap, key) {
  return colorMap.get(key) || CONFIG.COLORS.DEFAULT_BACKGROUND;
}

const DATE_FORMATS = {
  DATE_ONLY: "yyyy/MM/dd",
  DATETIME: "yyyy-MM-dd HH:mm",
  MONTH_DAY: "M/d",
  BACKUP_TIMESTAMP: "yyyy-MM-dd_HH-mm"
};

const BATCH_CONFIG = {
  CACHE_EXPIRATION: 300 // 5分
};/**
 * Config.gs
 * システム全体の設定定数を管理
 */

// =================================================================================
// === グローバル設定 ===
// =================================================================================
const CONFIG = {
  // Google DriveフォルダID
  FOLDERS: {
    REFERENCE_MATERIAL_PARENT: "124OR71hkr2jeT-5esv0GHAeZn83fAvYc",
    SERIES_MODEL_PARENT: "1XdiYBWiixF_zOSScT7UKUhCQkye3MLNJ",
    BACKUP_PARENT: "1HCDyjF_Kw2jlzN491uZK1X6QeXFuOWSl"
  },

  // シート名
  SHEETS: {
    MAIN: "メインシート",
    INPUT_PREFIX: "工数_",
    TANTOUSHA_MASTER: "担当者マスタ", // ★担当者マスタシートを定義
  },

  // 色設定
  COLORS: {
    DUPLICATE_HIGHLIGHT: '#e8eaed',
    DEFAULT_BACKGROUND: '#ffffff',
    WEEKEND_HOLIDAY: '#e8eaed'
  },

  // バックアップ設定
  BACKUP: {
    KEEP_COUNT: 5,
    PREFIX: "【Backup】"
  },

  // 日本の祝日カレンダーID
  HOLIDAY_CALENDAR_ID: 'ja.japanese#holiday@group.v.calendar.google.com',

  // デフォルト値
  DEFAULTS: {
    PROGRESS: "未着手",
    DUPLICATE_TEXT: "機番重複"
  }
};

// =================================================================================
// === 動的な色設定 ===
// =================================================================================
const PROGRESS_COLORS = new Map([
  ["未着手", "#ffcfc9"],
  ["仕掛中", "#e6cff2"],
  ["保留", "#c6dbe1"],
  ["完了", "#d4edbc"],
  ["機番重複", "#e8eaed"]
]);

const TANTOUSHA_COLORS = new Map([
  ["志賀", "#ffcfc9"],
  ["遠藤", "#d4edbc"],
  ["小板橋", "#bfe1f6"]
]);

const TOIAWASE_COLORS = new Map([
  ["問合済", "#ffcfc9"],
  ["回答済", "#bfe1f6"]
]);

// =================================================================================
// === シート列定義（ヘッダー名ベース） ===
// =================================================================================
/**
 * ★★★ 要望を反映した最終的な列構成 ★★★
 */
const MAIN_SHEET_HEADERS = {
  MGMT_NO: "管理No",
  KIBAN: "機番",
  MODEL: "機種",
  KIBAN_URL: "機番(リンク)",
  SERIES_URL: "STD資料(リンク)",
  REFERENCE_KIBAN: "参考製番",
  CIRCUIT_DIAGRAM_NO: "回路図番",
  TOIAWASE: "問い合わせ",
  TEMP_CODE: "仮コード",
  TANTOUSHA: "担当者",
  DESTINATION: "納入先",
  PLANNED_HOURS_PANEL: "予定工数(盤配)",
  ACTUAL_HOURS_PANEL: "実績工数(盤配)",
  PLANNED_HOURS_WIRE: "予定工数(線加工)",
  ACTUAL_HOURS_WIRE: "実績工数(線加工)",
  PROGRESS_PANEL: "進捗(盤配)",
  PROGRESS_WIRE: "進捗(線加工)",
  START_DATE: "仕掛日",
  COMPLETE_DATE: "完了日",
  DRAWING_DEADLINE: "作図期限",
  PROGRESS_EDITOR: "進捗記入者",
  UPDATE_TS: "更新日時",
  OVERRUN_REASON: "係り超過理由",
  NOTES: "注意点",
  REMARKS: "備考"
};

/**
 * 工数シートの列定義（メインシートと完全に一致させる）
 */
const INPUT_SHEET_HEADERS = MAIN_SHEET_HEADERS;


// =================================================================================
// === ユーティリティ関数（Config内） ===
// =================================================================================
/**
 * ヘッダー行から列インデックスを動的に取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {Object} headerDef - ヘッダー定義オブジェクト
 * @return {Object} ヘッダー名をキー、列インデックスを値とするオブジェクト
 */
function getColumnIndices(sheet, headerDef) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `col_indices_${sheet.getId()}_${sheet.getName()}`;
  const cached = cache.get(cacheKey);

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e) {}
  }

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const indices = {};
  for (const [key, headerName] of Object.entries(headerDef)) {
    const index = headerRow.indexOf(headerName) + 1;
    if (index > 0) {
      indices[key] = index;
    } else {
      Logger.log(`警告: ヘッダー "${headerName}" がシート "${sheet.getName()}" に見つかりません。`);
    }
  }

  cache.put(cacheKey, JSON.stringify(indices), 21600); // 6時間キャッシュ
  return indices;
}

/**
 * 色を取得する汎用関数
 */
function getColor(colorMap, key) {
  return colorMap.get(key) || CONFIG.COLORS.DEFAULT_BACKGROUND;
}

const DATE_FORMATS = {
  DATE_ONLY: "yyyy/MM/dd",
  DATETIME: "yyyy-MM-dd HH:mm",
  MONTH_DAY: "M/d",
  BACKUP_TIMESTAMP: "yyyy-MM-dd_HH-mm"
};

const BATCH_CONFIG = {
  CACHE_EXPIRATION: 300 // 5分
};/**
 * Config.gs
 * システム全体の設定定数を管理
 * 列の追加やステータスの追加に対応できる柔軟な設計
 */

// =================================================================================
// === グローバル設定 ===
// =================================================================================

const CONFIG = {
  // Google DriveフォルダID
  FOLDERS: {
    REFERENCE_MATERIAL_PARENT: "124OR71hkr2jeT-5esv0GHAeZn83fAvYc",
    SERIES_MODEL_PARENT: "1XdiYBWiixF_zOSScT7UKUhCQkye3MLNJ",
    BACKUP_PARENT: "1HCDyjF_Kw2jlzN491uZK1X6QeXFuOWSl"
  },

  // シート名
  SHEETS: {
    KIBAN_MASTER: "機番マスタ",
    MAIN: "メインシート",
    INPUT_PREFIX: "工数_",
    PRODUCTION_MASTER: "生産管理マスタ"
  },

  // 色設定
  COLORS: {
    DUPLICATE_HIGHLIGHT: '#e8eaed',
    DEFAULT_BACKGROUND: '#ffffff',
    WEEKEND_HOLIDAY: '#e8eaed'
  },

  // バックアップ設定
  BACKUP: {
    KEEP_COUNT: 5,
    PREFIX: "【Backup】"
  },

  // 日本の祝日カレンダーID
  HOLIDAY_CALENDAR_ID: 'ja.japanese#holiday@group.v.calendar.google.com',

  // デフォルト値
  DEFAULTS: {
    PROGRESS: "未着手",
    DUPLICATE_TEXT: "機番重複"
  }
};

// =================================================================================
// === 動的な色設定（ステータス追加に対応） ===
// =================================================================================

/**
 * 進捗ステータスと色のマッピング
 * 新しいステータスはここに追加するだけでOK
 */
const PROGRESS_COLORS = new Map([
  ["未着手", "#ffcfc9"],
  ["配置済", "#d4edbc"],
  ["ACS済", "#bfe1f6"],
  ["日精済", "#ffe5a0"],
  ["係り中", "#e6cff2"],
  ["機番重複", "#e8eaed"],
  ["保留", "#c6dbe1"],
  ["完了", "#d4edbc"]
  // 新しいステータスをここに追加
  // ["新ステータス", "#色コード"],
]);

/**
 * 担当者と色のマッピング
 */
const TANTOUSHA_COLORS = new Map([
  ["志賀", "#ffcfc9"],
  ["遠藤", "#d4edbc"],
  ["小板橋", "#bfe1f6"]
  // 新しい担当者をここに追加
]);

/**
 * 問い合わせステータスと色のマッピング
 */
const TOIAWASE_COLORS = new Map([
  ["問合済", "#ffcfc9"],
  ["回答済", "#bfe1f6"]
  // 新しいステータスをここに追加
]);

// =================================================================================
// === シート列定義（ヘッダー名ベース） ===
// =================================================================================

/**
 * 機番マスタシートの列定義
 * ヘッダー名をキーとして使用（列追加に対応）
 */
const KIBAN_MASTER_HEADERS = {
  MGMT_NO: "管理Ｎｏ．",
  KIBAN: "機番",
  MODEL: "機種",
  DESTINATION: "納入先",
  PLANNED_HOURS: "予定工数(h)",
  DRAWING_DEADLINE: "作図期限",
  PROGRESS: "進捗",
  FOLDER_URL: "製番資料",
  SERIES_FOLDER_URL: "STD資料"
  // 新しい列はここに追加
};

/**
 * メインシートの列定義
 */
const MAIN_SHEET_HEADERS = {
  MGMT_NO: "管理No",
  KIBAN: "機番",
  MODEL: "機種",
  KIBAN_URL: "機番(リンク)",
  SERIES_URL: "STD資料(リンク)",
  REFERENCE_KIBAN: "参考製番",
  TOIAWASE: "問い合わせ",
  TEMP_CODE: "仮コード",
  TANTOUSHA: "担当者",
  DESTINATION: "納入先",
  PLANNED_HOURS: "予定工数",
  TOTAL_LABOR: "合計工数",
  PROGRESS: "進捗",
  DRAWING_DEADLINE: "作図期限",
  PROGRESS_EDITOR: "進捗記入者",
  UPDATE_TS: "更新日時",
  COMPLETE_DATE: "完了日",
  ASSEMBLY_START: "組み立て開始日",
  REMARKS: "備考"
  // 新しい列はここに追加
};

/**
 * 工数シートの列定義
 */
const INPUT_SHEET_HEADERS = {
  MGMT_NO: "管理No.",
  KIBAN: "機番",
  MODEL: "機種",
  TANTOU: "担当者",
  TOIAWASE: "問合せ",
  DEADLINE: "作図期限",
  PROGRESS: "進捗",
  TIMESTAMP: "更新日時",
  PLANNED_HOURS: "予定工数",
  TOTAL_HOURS: "合計工数"
  // LABOR_START以降は日付列なので動的に処理
};

/**
 * 生産管理マスタの列定義
 */
const PROD_MASTER_HEADERS = {
  NO: "No",
  SEIBAN: "製番",
  MODEL: "機種",
  DESTINATION: "納入先",
  ASSEMBLY_START: "組立開始日"
  // 新しい列はここに追加
};

// =================================================================================
// === 列インデックス取得関数 ===
// =================================================================================

/**
 * ヘッダー行から列インデックスを動的に取得
 * @param {Sheet} sheet - 対象シート
 * @param {Object} headerDef - ヘッダー定義オブジェクト
 * @return {Object} ヘッダー名をキー、列インデックスを値とするオブジェクト
 */
function getColumnIndices(sheet, headerDef) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `col_indices_${sheet.getName()}`;
  const cached = cache.get(cacheKey);
  
  if (cached) {
    return JSON.parse(cached);
  }
  
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const indices = {};
  
  // ヘッダー定義のキーと実際のヘッダーをマッピング
  for (const [key, headerName] of Object.entries(headerDef)) {
    const index = headerRow.indexOf(headerName) + 1; // 1-based index
    if (index > 0) {
      indices[key] = index;
    } else {
      console.warn(`ヘッダー "${headerName}" が見つかりません: ${sheet.getName()}`);
    }
  }
  
  // 追加の列（定義外）も検出
  headerRow.forEach((header, idx) => {
    if (header && !Object.values(headerDef).includes(header)) {
      const key = header.replace(/[^\w]/g, '_').toUpperCase();
      indices[key] = idx + 1;
      console.log(`新しい列を検出: "${header}" (${key})`);
    }
  });
  
  cache.put(cacheKey, JSON.stringify(indices), BATCH_CONFIG.CACHE_EXPIRATION);
  return indices;
}

/**
 * 色を取得する汎用関数
 * @param {Map} colorMap - 色のマッピング
 * @param {string} key - ステータスや名前
 * @return {string} 色コード（見つからない場合はデフォルト色）
 */
function getColor(colorMap, key) {
  return colorMap.get(key) || CONFIG.COLORS.DEFAULT_BACKGROUND;
}

// =================================================================================
// === 日付フォーマット定義 ===
// =================================================================================

const DATE_FORMATS = {
  DATE_ONLY: "yyyy/MM/dd",
  DATETIME: "yyyy-MM-dd HH:mm",
  MONTH_DAY: "M/d",
  BACKUP_TIMESTAMP: "yyyy-MM-dd_HH-mm"
};

// =================================================================================
// === バッチ処理設定 ===
// =================================================================================

const BATCH_CONFIG = {
  MAX_ROWS_PER_BATCH: 500,
  MAX_CELLS_PER_REQUEST: 50000,
  CACHE_EXPIRATION: 300
};

// =================================================================================
// === エラーメッセージ ===
// =================================================================================

const ERROR_MESSAGES = {
  SHEET_NOT_FOUND: "エラー: 必要なシートが見つかりません。",
  NO_DATA: "情報: データがありません。",
  FOLDER_CREATE_ERROR: "フォルダ作成/取得エラー: ",
  BACKUP_ERROR: "バックアップエラー: ",
  COLUMN_NOT_FOUND: "エラー: 必要な列が見つかりません: "
};

// =================================================================================
// === 処理メッセージ ===
// =================================================================================

const PROCESS_MESSAGES = {
  START: "処理を開始します...",
  COMPLETE: "自動処理が完了しました。",
  UPDATE_COMPLETE: "更新が完了しました。",
  REBUILDING: "再構築中..."
};/**
 * Config.gs
 * システム全体の設定定数を管理
 */

// =================================================================================
// === グローバル設定 ===
// =================================================================================
const CONFIG = {
  // Google DriveフォルダID
  FOLDERS: {
    REFERENCE_MATERIAL_PARENT: "124OR71hkr2jeT-5esv0GHAeZn83fAvYc",
    SERIES_MODEL_PARENT: "1XdiYBWiixF_zOSScT7UKUhCQkye3MLNJ",
    BACKUP_PARENT: "1HCDyjF_Kw2jlzN491uZK1X6QeXFuOWSl"
  },

  // シート名
  SHEETS: {
    MAIN: "メインシート",
    INPUT_PREFIX: "工数_",
    TANTOUSHA_MASTER: "担当者マスタ", // ★担当者マスタシートを定義
  },

  // 色設定
  COLORS: {
    DUPLICATE_HIGHLIGHT: '#e8eaed',
    DEFAULT_BACKGROUND: '#ffffff',
    WEEKEND_HOLIDAY: '#e8eaed'
  },

  // バックアップ設定
  BACKUP: {
    KEEP_COUNT: 5,
    PREFIX: "【Backup】"
  },

  // 日本の祝日カレンダーID
  HOLIDAY_CALENDAR_ID: 'ja.japanese#holiday@group.v.calendar.google.com',

  // デフォルト値
  DEFAULTS: {
    PROGRESS: "未着手",
    DUPLICATE_TEXT: "機番重複"
  }
};

// =================================================================================
// === 動的な色設定 ===
// =================================================================================
const PROGRESS_COLORS = new Map([
  ["未着手", "#ffcfc9"],
  ["仕掛中", "#e6cff2"],
  ["保留", "#c6dbe1"],
  ["完了", "#d4edbc"],
  ["機番重複", "#e8eaed"]
]);

const TANTOUSHA_COLORS = new Map([
  ["志賀", "#ffcfc9"],
  ["遠藤", "#d4edbc"],
  ["小板橋", "#bfe1f6"]
]);

const TOIAWASE_COLORS = new Map([
  ["問合済", "#ffcfc9"],
  ["回答済", "#bfe1f6"]
]);

// =================================================================================
// === シート列定義（ヘッダー名ベース） ===
// =================================================================================
/**
 * ★★★ 要望を反映した最終的な列構成 ★★★
 */
const MAIN_SHEET_HEADERS = {
  MGMT_NO: "管理No",
  KIBAN: "機番",
  MODEL: "機種",
  KIBAN_URL: "機番(リンク)",
  SERIES_URL: "STD資料(リンク)",
  REFERENCE_KIBAN: "参考製番",
  CIRCUIT_DIAGRAM_NO: "回路図番",
  TOIAWASE: "問い合わせ",
  TEMP_CODE: "仮コード",
  TANTOUSHA: "担当者",
  DESTINATION: "納入先",
  PLANNED_HOURS_PANEL: "予定工数(盤配)",
  ACTUAL_HOURS_PANEL: "実績工数(盤配)",
  PLANNED_HOURS_WIRE: "予定工数(線加工)",
  ACTUAL_HOURS_WIRE: "実績工数(線加工)",
  PROGRESS_PANEL: "進捗(盤配)",
  PROGRESS_WIRE: "進捗(線加工)",
  START_DATE: "仕掛日",
  COMPLETE_DATE: "完了日",
  DRAWING_DEADLINE: "作図期限",
  PROGRESS_EDITOR: "進捗記入者",
  UPDATE_TS: "更新日時",
  OVERRUN_REASON: "係り超過理由",
  NOTES: "注意点",
  REMARKS: "備考"
};

/**
 * 工数シートの列定義（メインシートと完全に一致させる）
 */
const INPUT_SHEET_HEADERS = MAIN_SHEET_HEADERS;


// =================================================================================
// === ユーティリティ関数（Config内） ===
// =================================================================================
/**
 * ヘッダー行から列インデックスを動的に取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {Object} headerDef - ヘッダー定義オブジェクト
 * @return {Object} ヘッダー名をキー、列インデックスを値とするオブジェクト
 */
function getColumnIndices(sheet, headerDef) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `col_indices_${sheet.getId()}_${sheet.getName()}`;
  const cached = cache.get(cacheKey);

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e) {}
  }

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const indices = {};
  for (const [key, headerName] of Object.entries(headerDef)) {
    const index = headerRow.indexOf(headerName) + 1;
    if (index > 0) {
      indices[key] = index;
    } else {
      Logger.log(`警告: ヘッダー "${headerName}" がシート "${sheet.getName()}" に見つかりません。`);
    }
  }

  cache.put(cacheKey, JSON.stringify(indices), 21600); // 6時間キャッシュ
  return indices;
}

/**
 * 色を取得する汎用関数
 */
function getColor(colorMap, key) {
  return colorMap.get(key) || CONFIG.COLORS.DEFAULT_BACKGROUND;
}

const DATE_FORMATS = {
  DATE_ONLY: "yyyy/MM/dd",
  DATETIME: "yyyy-MM-dd HH:mm",
  MONTH_DAY: "M/d",
  BACKUP_TIMESTAMP: "yyyy-MM-dd_HH-mm"
};

const BATCH_CONFIG = {
  CACHE_EXPIRATION: 300 // 5分
};