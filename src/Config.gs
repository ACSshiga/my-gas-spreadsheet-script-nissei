/**
 * Config.gs
 * システム全体の設定定数を管理します。
 */
const CONFIG = {
  FOLDERS: {
    REFERENCE_MATERIAL_PARENT: "124OR71hkr2jeT-5esv0GHAeZn83fAvYc", // 機番フォルダの親フォルダID
    SERIES_MODEL_PARENT:       "1XdiYBWiixF_zOSScT7UKUhCQkye3MLNJ", // STD資料フォルダの親フォルダID
    BACKUP_PARENT:             "1OKyXDvCMDiAsvcZXac2BjuJDk5x1-JyO",
    IMPORT_SOURCE_FOLDER: "16RqfEDRYnoq09Xu2mmM3e8wQbqulOh56", // 申請書インポート用
    PROCESSED_FOLDER:     "1T5W9vYZKoZDtD-6_clPGRUolgqpK4Zgg"  // 処理済み申請書
  },
  // ▼▼▼▼▼ このセクションを新しく追加 ▼▼▼▼▼
  TEMPLATES: {
    // 顧客用管理表のテンプレート（スプレッドシート）のID
    MANAGEMENT_SHEET_TEMPLATE_ID: "1BXLGtgjQelZVOUpoRrbM0bUqjhNlA1TW" 
  },
  // ▲▲▲▲▲ ここまで追加 ▲▲▲▲▲
  SHEETS: {
    MAIN:             "メインシート",
    INPUT_PREFIX:     "工数_",
    BILLING:          "請求シート",
    TANTOUSHA_MASTER: "担当者マスタ",
    SAGYOU_KUBUN_MASTER: "作業区分マスタ",
    SHINCHOKU_MASTER: "進捗マスタ",
    TOIAWASE_MASTER:  "問い合わせマスタ"
  },
  COLORS: {
    DEFAULT_BACKGROUND: '#ffffff',
    ALTERNATE_ROW:      '#f0f5f5', // おしゃれなミントグリーン
    HEADER_BACKGROUND:  '#4f5459', // ヘッダー用の背景色
    WEEKEND_HOLIDAY:    '#f2f2f2'
  },
  DATA_START_ROW: {
    MAIN: 2,
    INPUT: 3
  },
  BACKUP: {
    KEEP_COUNT: 5,
    PREFIX:     "【Backup】"
  },
  HOLIDAY_CALENDAR_ID: 'ja.japanese#holiday@group.v.calendar.google.com',
};

const MAIN_SHEET_HEADERS = {
  MGMT_NO: "管理No", SAGYOU_KUBUN: "作業区分", KIBAN: "機番", MODEL: "機種",
  KIBAN_URL: "機番(リンク)", SERIES_URL: "STD資料(リンク)", REFERENCE_KIBAN: "参考製番",
  CIRCUIT_DIAGRAM_NO: "回路図番", TOIAWASE: "問い合わせ", TEMP_CODE: "仮コード",
  TANTOUSHA: "担当者", DESTINATION: "納入先", PLANNED_HOURS: "予定工数",
  ACTUAL_HOURS: "実績工数", PROGRESS: "進捗", START_DATE: "仕掛日",
  COMPLETE_DATE: "完了日", DRAWING_DEADLINE: "作図期限", PROGRESS_EDITOR: "進捗記入者",
  UPDATE_TS: "更新日時", OVERRUN_REASON: "係り超過理由", NOTES: "注意点", REMARKS: "備考"
};

const INPUT_SHEET_HEADERS = {
  MGMT_NO: "管理No", SAGYOU_KUBUN: "作業区分", KIBAN: "機番", PROGRESS: "進捗",
  PLANNED_HOURS:"予定工数", ACTUAL_HOURS_SUM: "実績工数合計", SEPARATOR: "",
};

function getColumnIndices(sheet, headerDef) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `col_indices_${sheet.getSheetId()}`;
  const cached = cache.get(cacheKey);
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const indices = {};
  for (const [key, headerName] of Object.entries(headerDef)) {
    const index = headerRow.indexOf(headerName) + 1;
    if (index > 0) indices[key] = index;
  }
  cache.put(cacheKey, JSON.stringify(indices), 21600);
  return indices;
}

function getColor(colorMap, key, defaultValue = CONFIG.COLORS.DEFAULT_BACKGROUND) {
  return colorMap.get(key) || defaultValue;
}

const DATE_FORMATS = {
  DATE_ONLY: "yyyy/MM/dd", DATETIME: "yyyy-MM-dd HH:mm:ss",
  MONTH_DAY: "M/d", BACKUP_TIMESTAMP: "yyyy-MM-dd_HH-mm"
};