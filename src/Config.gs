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
    REFERENCE_MATERIAL_PARENT: "1FVzgvod5z9jdbI0yrQo8G8cH0Bv1a1pn", // 機番(リンク)の親フォルダ
    SERIES_MODEL_PARENT:       "1Qol5fRzYxvEfzo9BkNVmjnGELrkgvVel", // STD資料(リンク)の親フォルダ
    BACKUP_PARENT:             "1OKyXDvCMDiAsvcZXac2BjuJDk5x1-JyO"  // バックアップの親フォルダ
  },

  // シート名
  SHEETS: {
    MAIN:             "メインシート",
    INPUT_PREFIX:     "工数_",
    BILLING:          "請求シート",
    TANTOUSHA_MASTER: "担当者マスタ",
    SAGYOU_KUBUN_MASTER: "作業区分マスタ",
    SHINCHOKU_MASTER: "進捗マスタ",
    TOIAWASE_MASTER:  "問い合わせマスタ"
  },

  // 色設定
  COLORS: {
    DEFAULT_BACKGROUND: '#ffffff', // デフォルトの背景色
    WEEKEND_HOLIDAY:    '#f2f2f2', // 休日の背景色
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
    BILLING:          "請求シート", // ★請求シート名を追加
    TANTOUSHA_MASTER: "担当者マスタ",
    SAGYOU_KUBUN_MASTER: "作業区分マスタ",
    SHINCHOKU_MASTER: "進捗マスタ",
    TOIAWASE_MASTER:  "問い合わせマスタ"
  },

  // 色設定
  COLORS: {
    DEFAULT_BACKGROUND: '#ffffff', // デフォルトの背景色
    WEEKEND_HOLIDAY:    '#f2f2f2', // 休日の背景色
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
};