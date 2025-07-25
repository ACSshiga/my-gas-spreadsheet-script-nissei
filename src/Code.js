// =================================================================================
// === グローバル設定・ユーティリティ関数 (最終レイアウト対応版) ===
// =================================================================================

const DUPLICATE_HIGHLIGHT_COLOR = '#e8eaed';
const DEFAULT_CELL_BACKGROUND_COLOR = '#ffffff';

const REFERENCE_MATERIAL_PARENT_FOLDER_ID = "124OR71hkr2jeT-5esv0GHAeZn83fAvYc";
const SERIES_MODEL_PARENT_FOLDER_ID = "1XdiYBWiixF_zOSScT7UKUhCQkye3MLNJ";
const BACKUP_PARENT_FOLDER_ID = "1HCDyjF_Kw2jlzN491uZK1X6QeXFuOWSl";

// ★★★ 機番マスタ 定数 (最終レイアウト) ★★★
const KIBAN_MASTER_SHEET_NAME = "機番マスタ";
const KIBAN_MASTER_MGMT_NO_COL = 1;      // A列: 管理Ｎｏ．
const KIBAN_MASTER_KIBAN_COL = 2;        // B列: 機番
const KIBAN_MASTER_MODEL_COL = 3;        // C列: 機種
const KIBAN_MASTER_DESTINATION_COL = 4;  // D列: 納入先
const KIBAN_MASTER_PLANNED_HOURS_COL = 5;  // E列: 予定工数(h)
const KIBAN_MASTER_DRAWING_DEADLINE_COL = 6; // F列: 作図期限
const KIBAN_MASTER_PROGRESS_COL = 7;     // G列: 進捗
const KIBAN_MASTER_FOLDER_URL_COL = 8;     // H列: 製番資料
const KIBAN_MASTER_SERIES_FOLDER_URL_COL = 9; // I列: STD資料

// ★★★ メインシート 定数 (最終レイアウト) ★★★
const MAIN_SHEET_NAME = "メインシート";
const MAIN_SHEET_MGMT_NO_COL = 1;         // A列: 管理No
const MAIN_SHEET_KIBAN_COL = 2;           // B列: 機番(リンクなし)
const MAIN_SHEET_MODEL_COL = 3;           // C列: 機種(リンクなし)
const MAIN_SHEET_KIBAN_URL_COL = 4;         // D列: 機番(ハイパーリンクあり)
const MAIN_SHEET_SERIES_URL_COL = 5;      // E列: STD資料(ハイパーリンクあり)
const MAIN_SHEET_REFERENCE_KIBAN_COL = 6; // F列: 参考製番
const MAIN_SHEET_TOIAWASE_COL = 7;        // G列: 問い合わせ
const MAIN_SHEET_TEMP_CODE_COL = 8;         // H列: 仮コード
const MAIN_SHEET_TANTOUSHA_COL = 9;         // I列: 担当者
const MAIN_SHEET_DESTINATION_COL = 10;        // J列: 納入先
const MAIN_SHEET_PLANNED_HOURS_COL = 11;      // K列: 予定工数
const MAIN_SHEET_TOTAL_LABOR_COL = 12;      // L列: 合計工数
const MAIN_SHEET_PROGRESS_COL = 13;         // M列: 進捗
const MAIN_SHEET_DRAWING_DEADLINE_COL = 14;   // N列: 作図期限
const MAIN_SHEET_PROGRESS_EDITOR_COL = 15;    // O列: 進捗記入者
const MAIN_SHEET_UPDATE_TS_COL = 16;        // P列: 更新日時
const MAIN_SHEET_COMPLETE_DATE_COL = 17;      // Q列: 完了日
const MAIN_SHEET_ASSEMBLY_START_COL = 18;   // R列: 組み立て開始日
const MAIN_SHEET_REMARKS_COL = 19;        // S列: 備考

// ★★★ 工数シート 定数 (最終レイアウト) ★★★
const INPUT_SHEET_PREFIX = "工数_";
const INPUT_SHEET_MGMT_NO_COL = 1;       // A列: 管理No.
const INPUT_SHEET_KIBAN_COL = 2;       // B列: 機番
const INPUT_SHEET_MODEL_COL = 3;       // C列: 機種
const INPUT_SHEET_TANTOU_COL = 4;        // D列: 担当者
const INPUT_SHEET_TOIAWASE_COL = 5;      // E列: 問合せ
const INPUT_SHEET_DEADLINE_COL = 6;    // F列: 作図期限
const INPUT_SHEET_PROGRESS_COL = 7;    // G列: 進捗
const INPUT_SHEET_TIMESTAMP_COL = 8;       // H列: 更新日時
const INPUT_SHEET_PLANNED_HOURS_COL = 9;   // I列: 予定工数
const INPUT_SHEET_TOTAL_HOURS_COL = 10;      // J列: 合計工数
const INPUT_SHEET_LABOR_START_COL = 11;      // K列: 日付工数開始

function formatDateForComparison(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) { return null; }
  const year = date.getFullYear();
  const month = ('0' + (date.getMonth() + 1)).slice(-2);
  const day = ('0' + date.getDate()).slice(-2);
  return year + '-' + month + '-' + day;
}

// =================================================================================
// === メインシート関連処理 (ハイパーリンクエラー修正版) ===
// =================================================================================

function updateMainSheetFromMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(KIBAN_MASTER_SHEET_NAME);
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);

  if (!masterSheet || !mainSheet) {
    Logger.log("エラー: 必要なシートが見つかりません。");
    return;
  }

  const masterLastRow = masterSheet.getLastRow();
  if (masterLastRow <= 1) {
    Logger.log("情報: 「" + KIBAN_MASTER_SHEET_NAME + "」にデータがありません。");
    return;
  }
  
  const masterRange = masterSheet.getRange(2, 1, masterLastRow - 1, masterSheet.getLastColumn());
  const masterValues = masterRange.getValues();
  const masterKibanUrlRichTextValues = masterSheet.getRange(2, KIBAN_MASTER_FOLDER_URL_COL, masterLastRow - 1, 1).getRichTextValues();
  const masterSeriesUrlRichTextValues = masterSheet.getRange(2, KIBAN_MASTER_SERIES_FOLDER_URL_COL, masterLastRow - 1, 1).getRichTextValues();
  
  const masterMap = new Map();
  masterValues.forEach((row, index) => {
    const managementNo = String(row[KIBAN_MASTER_MGMT_NO_COL - 1]).trim();
    if (managementNo) {
      const kibanStr = String(row[KIBAN_MASTER_KIBAN_COL - 1]);
      const modelStr = String(row[KIBAN_MASTER_MODEL_COL - 1]);
      const seriesFolderNameForLabel = extractSeriesPlusInitialNumber_(modelStr) || kibanStr;

      let kibanUrl = "";
      const kibanRichText = masterKibanUrlRichTextValues[index][0];
      const kibanLinkUrl = kibanRichText.getLinkUrl();
      const kibanDisplayText = kibanRichText.getText() || kibanStr;

      if (kibanLinkUrl) {
        // ★★★ 修正箇所: URLと表示テキスト内の " を "" に変換してエスケープする ★★★
        const safeLinkUrl = kibanLinkUrl.replace(/"/g, '""');
        const safeDisplayText = kibanDisplayText.replace(/"/g, '""');
        kibanUrl = `=HYPERLINK("${safeLinkUrl}", "${safeDisplayText}")`;
      } else if (kibanDisplayText.toLowerCase().startsWith("http")) {
        const safeLinkUrl = kibanDisplayText.replace(/"/g, '""');
        const safeDisplayText = kibanStr.replace(/"/g, '""');
        kibanUrl = `=HYPERLINK("${safeLinkUrl}", "${safeDisplayText}")`;
      }
      
      let seriesUrl = "";
      const seriesRichText = masterSeriesUrlRichTextValues[index][0];
      const seriesLinkUrl = seriesRichText.getLinkUrl();
      const seriesDisplayText = seriesRichText.getText() || seriesFolderNameForLabel;
      
      if (seriesLinkUrl) {
        // ★★★ 修正箇所: こちらも同様に " をエスケープ ★★★
        const safeLinkUrl = seriesLinkUrl.replace(/"/g, '""');
        const safeDisplayText = seriesDisplayText.replace(/"/g, '""');
        seriesUrl = `=HYPERLINK("${safeLinkUrl}", "${safeDisplayText}")`;
      } else if (seriesDisplayText.toLowerCase().startsWith("http")) {
        const safeLinkUrl = seriesDisplayText.replace(/"/g, '""');
        const safeDisplayText = seriesFolderNameForLabel.replace(/"/g, '""');
        seriesUrl = `=HYPERLINK("${safeLinkUrl}", "${safeDisplayText}")`;
      }
      
      masterMap.set(managementNo, {
        kiban: kibanStr,
        model: modelStr,
        kibanUrl: kibanUrl,
        seriesUrl: seriesUrl,
        destination: String(row[KIBAN_MASTER_DESTINATION_COL - 1]),
        plannedHours: row[KIBAN_MASTER_PLANNED_HOURS_COL - 1],
        drawingDeadline: row[KIBAN_MASTER_DRAWING_DEADLINE_COL - 1],
        progress: String(row[KIBAN_MASTER_PROGRESS_COL - 1]),
      });
    }
  });

  const mainLastRow = mainSheet.getLastRow();
  const numCols = mainSheet.getMaxColumns();
  const oldRows = mainLastRow > 1 ? mainSheet.getRange(2, 1, mainLastRow - 1, numCols).getValues() : [];
  const oldRowsMap = new Map(oldRows.map(row => [String(row[MAIN_SHEET_MGMT_NO_COL - 1]).trim(), row]));
  const newRows = [];

  masterMap.forEach((info, managementNo) => {
    const existingRow = oldRowsMap.get(managementNo) || [];
    let outputRow = new Array(numCols).fill("");
    
    outputRow[MAIN_SHEET_MGMT_NO_COL - 1] = managementNo;
    outputRow[MAIN_SHEET_KIBAN_COL - 1] = info.kiban;
    outputRow[MAIN_SHEET_MODEL_COL - 1] = info.model;
    outputRow[MAIN_SHEET_KIBAN_URL_COL - 1] = info.kibanUrl;
    outputRow[MAIN_SHEET_SERIES_URL_COL - 1] = info.seriesUrl;
    outputRow[MAIN_SHEET_DESTINATION_COL - 1] = info.destination;
    outputRow[MAIN_SHEET_PLANNED_HOURS_COL - 1] = info.plannedHours;
    outputRow[MAIN_SHEET_PROGRESS_COL - 1] = info.progress;
    outputRow[MAIN_SHEET_DRAWING_DEADLINE_COL - 1] = info.drawingDeadline;
    
    if (existingRow.length > 0) {
      outputRow[MAIN_SHEET_REFERENCE_KIBAN_COL - 1] = existingRow[MAIN_SHEET_REFERENCE_KIBAN_COL - 1];
      outputRow[MAIN_SHEET_TOIAWASE_COL - 1] = existingRow[MAIN_SHEET_TOIAWASE_COL - 1];
      outputRow[MAIN_SHEET_TEMP_CODE_COL - 1] = existingRow[MAIN_SHEET_TEMP_CODE_COL - 1];
      outputRow[MAIN_SHEET_TANTOUSHA_COL - 1] = existingRow[MAIN_SHEET_TANTOUSHA_COL - 1];
      outputRow[MAIN_SHEET_TOTAL_LABOR_COL - 1] = existingRow[MAIN_SHEET_TOTAL_LABOR_COL - 1];
      outputRow[MAIN_SHEET_PROGRESS_EDITOR_COL - 1] = existingRow[MAIN_SHEET_PROGRESS_EDITOR_COL - 1];
      outputRow[MAIN_SHEET_UPDATE_TS_COL - 1] = existingRow[MAIN_SHEET_UPDATE_TS_COL - 1];
      outputRow[MAIN_SHEET_COMPLETE_DATE_COL - 1] = existingRow[MAIN_SHEET_COMPLETE_DATE_COL - 1];
      outputRow[MAIN_SHEET_ASSEMBLY_START_COL - 1] = existingRow[MAIN_SHEET_ASSEMBLY_START_COL - 1];
      outputRow[MAIN_SHEET_REMARKS_COL - 1] = existingRow[MAIN_SHEET_REMARKS_COL - 1];
    }
    newRows.push(outputRow);
  });

  if (mainLastRow > 1) {
    mainSheet.getRange(2, 1, mainLastRow - 1, numCols).clearContent();
  }
  
  if (newRows.length > 0) {
    mainSheet.getRange(2, 1, newRows.length, numCols).setValues(newRows);
    mainSheet.getRange(2, MAIN_SHEET_DRAWING_DEADLINE_COL, newRows.length, 1).setNumberFormat("yyyy/MM/dd");
    mainSheet.getRange(2, MAIN_SHEET_UPDATE_TS_COL, newRows.length, 1).setNumberFormat("yyyy-MM-dd HH:mm");
    mainSheet.getRange(2, MAIN_SHEET_ASSEMBLY_START_COL, newRows.length, 1).setNumberFormat("yyyy/MM/dd");
    Logger.log(MAIN_SHEET_NAME + "を機番マスタに基づいて更新しました。");
  }
}
function applyProgressPlaceholder() {
  const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!mainSheet) return;
  const lastRow = mainSheet.getLastRow();
  if (lastRow <= 1) return;
  const range = mainSheet.getRange(2, MAIN_SHEET_PROGRESS_COL, lastRow - 1, 1);
  const values = range.getValues();
  const newValues = values.map(row => [row[0] || "未着手"]);
  if (JSON.stringify(values) !== JSON.stringify(newValues)) {
    range.setValues(newValues);
  }
}

function batchUpdateCompletionDates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  const range = sheet.getRange(2, MAIN_SHEET_PROGRESS_COL, lastRow - 1, MAIN_SHEET_COMPLETE_DATE_COL - MAIN_SHEET_PROGRESS_COL + 1);
  const values = range.getValues();
  const today = new Date();
  const newDates = values.map(row => {
    const progress = row[0];
    const completeDate = row[MAIN_SHEET_COMPLETE_DATE_COL - MAIN_SHEET_PROGRESS_COL];
    if ((progress === "完了" || progress === "ACS済" || progress === "日精済") && !completeDate) {
      return [today];
    }
    return [completeDate];
  });
  sheet.getRange(2, MAIN_SHEET_COMPLETE_DATE_COL, newDates.length, 1)
       .setValues(newDates)
       .setNumberFormat('yyyy-MM-dd');
}

function updateMainSheetLaborTotal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!mainSheet) return;
  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(INPUT_SHEET_PREFIX));
  const mainLastRow = mainSheet.getLastRow();
  if (mainLastRow <= 1) return;
  const mainMgmtNos = mainSheet.getRange(2, MAIN_SHEET_MGMT_NO_COL, mainLastRow - 1, 1).getValues().map(row => String(row[0]).trim());
  const laborTotals = new Map(mainMgmtNos.map(no => [no, 0]));

  inputSheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    const data = sheet.getRange(3, 1, lastRow - 2, INPUT_SHEET_TOTAL_HOURS_COL).getValues();
    data.forEach(row => {
      const mgmtNo = String(row[INPUT_SHEET_MGMT_NO_COL - 1]).trim();
      const labor = Number(row[INPUT_SHEET_TOTAL_HOURS_COL - 1]) || 0;
      if (laborTotals.has(mgmtNo)) {
        laborTotals.set(mgmtNo, laborTotals.get(mgmtNo) + labor);
      }
    });
  });

  const newTotalLaborValues = mainMgmtNos.map(no => [laborTotals.get(no) || 0]);
  mainSheet.getRange(2, MAIN_SHEET_TOTAL_LABOR_COL, newTotalLaborValues.length, 1).setValues(newTotalLaborValues);
}

function syncAssemblyStartDateFromProdMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productionMasterSheet = ss.getSheetByName("生産管理マスタ");
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);

  if (!productionMasterSheet || !mainSheet) return;
  const prodMasterLastRow = productionMasterSheet.getLastRow();
  if (prodMasterLastRow <= 1) return;
  const productionMasterValues = productionMasterSheet.getRange(2, 1, prodMasterLastRow - 1, 5).getValues();
  const productionDataMap = new Map();
  productionMasterValues.forEach(row => {
    const seiban = String(row[1]).trim();
    const assemblyStartDate = row[4];
    if (seiban && assemblyStartDate) {
      productionDataMap.set(seiban, assemblyStartDate);
    }
  });

  const mainSheetLastRow = mainSheet.getLastRow();
  if (mainSheetLastRow <= 1) return;
  const mainSheetRange = mainSheet.getRange(2, 1, mainSheetLastRow - 1, MAIN_SHEET_ASSEMBLY_START_COL);
  const mainSheetValues = mainSheetRange.getValues();
  let updated = false;
  
  const newValues = mainSheetValues.map(row => {
    const kiban = String(row[MAIN_SHEET_KIBAN_COL - 1]).trim();
    if (productionDataMap.has(kiban)) {
      const masterDate = new Date(productionDataMap.get(kiban));
      const currentDate = row[MAIN_SHEET_ASSEMBLY_START_COL - 1] ? new Date(row[MAIN_SHEET_ASSEMBLY_START_COL - 1]) : null;
      if (!currentDate || currentDate.getTime() !== masterDate.getTime()) {
        row[MAIN_SHEET_ASSEMBLY_START_COL - 1] = masterDate;
        updated = true;
      }
    }
    return row;
  });

  if (updated) {
    mainSheetRange.setValues(newValues);
  }
}

// =================================================================================
// === 工数シート関連処理 (最終レイアウト対応版) ===
// =================================================================================
function rebuildInputSheetsFromMainOptimized() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!mainSheet) {
    Logger.log("エラー: シート「" + MAIN_SHEET_NAME + "」が見つかりません。");
    return;
  }
  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(INPUT_SHEET_PREFIX));
  const mainLastRow = mainSheet.getLastRow();
  if (mainLastRow <= 1) {
    Logger.log("情報: 「" + MAIN_SHEET_NAME + "」にデータがありません。");
    return;
  }
  
  const numColsToReadFromMain = Math.max(
    MAIN_SHEET_KIBAN_URL_COL,
    MAIN_SHEET_SERIES_URL_COL,
    MAIN_SHEET_TANTOUSHA_COL,
    MAIN_SHEET_TOIAWASE_COL,
    MAIN_SHEET_DRAWING_DEADLINE_COL,
    MAIN_SHEET_PLANNED_HOURS_COL
  );
  const mainRange = mainSheet.getRange(2, 1, mainLastRow - 1, numColsToReadFromMain);
  const mainValues = mainRange.getValues();
  const mainFormulas = mainRange.getFormulas();

  const mainMap = new Map();
  mainValues.forEach((row, index) => {
    const mgmtNo = String(row[MAIN_SHEET_MGMT_NO_COL - 1]).trim();
    if (mgmtNo) {
      mainMap.set(mgmtNo, {
        kiban: row[MAIN_SHEET_KIBAN_COL - 1],
        model: row[MAIN_SHEET_MODEL_COL - 1],
        kibanUrl: mainFormulas[index][MAIN_SHEET_KIBAN_URL_COL - 1],
        seriesUrl: mainFormulas[index][MAIN_SHEET_SERIES_URL_COL - 1],
        tantousha: row[MAIN_SHEET_TANTOUSHA_COL - 1],
        toiawase: row[MAIN_SHEET_TOIAWASE_COL - 1],
        drawingDeadline: row[MAIN_SHEET_DRAWING_DEADLINE_COL - 1],
        plannedHours: row[MAIN_SHEET_PLANNED_HOURS_COL - 1],
      });
    }
  });

  inputSheets.forEach(sheet => {
    const sheetLastRow = sheet.getLastRow();
    const dataStartRow = 3;
    const maxCols = Math.max(INPUT_SHEET_LABOR_START_COL - 1, sheet.getLastColumn());
    
    let existingValues = [];
    if (sheetLastRow >= dataStartRow) {
      existingValues = sheet.getRange(dataStartRow, 1, sheetLastRow - dataStartRow + 1, maxCols).getValues();
    }
    const existingMap = new Map(existingValues.map(row => [String(row[INPUT_SHEET_MGMT_NO_COL - 1]).trim(), row]));

    const newSheetRowsData = [];
    mainMap.forEach((mainInfo, mgmtNo) => {
      const existingRow = existingMap.get(mgmtNo) || new Array(maxCols).fill("");
      let outputRow = [...existingRow];
      
      outputRow[INPUT_SHEET_MGMT_NO_COL - 1] = mgmtNo;
      
      if (mainInfo.kibanUrl && String(mainInfo.kibanUrl).startsWith('=')) {
        outputRow[INPUT_SHEET_KIBAN_COL - 1] = mainInfo.kibanUrl;
      } else {
        outputRow[INPUT_SHEET_KIBAN_COL - 1] = mainInfo.kiban;
      }

      if (mainInfo.seriesUrl && String(mainInfo.seriesUrl).startsWith('=')) {
        outputRow[INPUT_SHEET_MODEL_COL - 1] = mainInfo.seriesUrl;
      } else {
        outputRow[INPUT_SHEET_MODEL_COL - 1] = mainInfo.model;
      }

      outputRow[INPUT_SHEET_TANTOU_COL - 1] = mainInfo.tantousha;
      outputRow[INPUT_SHEET_TOIAWASE_COL - 1] = mainInfo.toiawase;
      outputRow[INPUT_SHEET_DEADLINE_COL - 1] = mainInfo.drawingDeadline;
      outputRow[INPUT_SHEET_PLANNED_HOURS_COL - 1] = mainInfo.plannedHours;
      outputRow[INPUT_SHEET_TOTAL_HOURS_COL - 1] = ""; 
      
      newSheetRowsData.push(outputRow);
    });

    if (sheetLastRow >= dataStartRow) {
      sheet.getRange(dataStartRow, 1, sheetLastRow - dataStartRow + 1, maxCols).clearContent();
    }

    if (newSheetRowsData.length > 0) {
      const range = sheet.getRange(dataStartRow, 1, newSheetRowsData.length, maxCols);
      range.setValues(newSheetRowsData);
      
      const sumFormulasR1C1 = newSheetRowsData.map(() => [`=IFERROR(SUM(RC[1]:RC[999]), 0)`]);
      sheet.getRange(dataStartRow, INPUT_SHEET_TOTAL_HOURS_COL, newSheetRowsData.length, 1).setFormulasR1C1(sumFormulasR1C1);
      
      sheet.getRange(dataStartRow, INPUT_SHEET_DEADLINE_COL, newSheetRowsData.length, 1).setNumberFormat('yyyy/MM/dd');
      sheet.getRange(dataStartRow, INPUT_SHEET_PLANNED_HOURS_COL, newSheetRowsData.length, 2).setNumberFormat('0.0');
    }
    Logger.log(`シート「${sheet.getName()}」を再構築しました。`);
  });
}

function markOrphanedRowsInInputSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!mainSheet) return;
  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(INPUT_SHEET_PREFIX));
  const mainLastRow = mainSheet.getLastRow();
  if (mainLastRow <= 1) return;
  const mainMgmtNos = mainSheet.getRange(2, 1, mainLastRow - 1, 1).getValues().map(row => String(row[0]).trim());
  const mainSet = new Set(mainMgmtNos);

  inputSheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    const range = sheet.getRange(3, 1, lastRow - 2, 1);
    const values = range.getValues();
    const backgrounds = range.getBackgrounds();
    const newBgs = [];
    let updated = false;

    values.forEach((val, i) => {
      const mgmtNo = String(val[0]).trim();
      const currentBg = backgrounds[i][0];
      let targetBg = DEFAULT_CELL_BACKGROUND_COLOR;
      if (!mgmtNo || !mainSet.has(mgmtNo)) {
        targetBg = DUPLICATE_HIGHLIGHT_COLOR;
      }
      if (currentBg.toLowerCase() !== targetBg.toLowerCase()) {
        updated = true;
      }
      newBgs.push([targetBg]);
    });

    if (updated) {
      range.setBackgrounds(newBgs);
    }
  });
}

function syncProgressFromInputToMain() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!mainSheet) return;
  const inputSheets = ss.getSheets().filter(s => s.getName().startsWith(INPUT_SHEET_PREFIX));
  const latestProgressMap = new Map();

  inputSheets.forEach(sheet => {
    const editorName = sheet.getName().replace(INPUT_SHEET_PREFIX, '');
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    const data = sheet.getRange(3, 1, lastRow - 2, INPUT_SHEET_TIMESTAMP_COL).getValues();
    data.forEach(row => {
      const mgmtNo = String(row[INPUT_SHEET_MGMT_NO_COL - 1]).trim();
      const progress = row[INPUT_SHEET_PROGRESS_COL - 1];
      const timestamp = row[INPUT_SHEET_TIMESTAMP_COL - 1];
      if (mgmtNo && progress && timestamp) {
        const newTimestamp = new Date(timestamp);
        if (isNaN(newTimestamp.getTime())) return;
        const existing = latestProgressMap.get(mgmtNo);
        if (!existing || newTimestamp > existing.timestamp) {
          latestProgressMap.set(mgmtNo, { progress, editorName, timestamp: newTimestamp });
        }
      }
    });
  });

  const mainLastRow = mainSheet.getLastRow();
  if (mainLastRow <= 1) return;

  const mainMgmtNos = mainSheet.getRange(2, MAIN_SHEET_MGMT_NO_COL, mainLastRow - 1, 1).getValues();
  
  let progressUpdates = [];
  let editorUpdates = [];
  let timestampUpdates = [];

  mainMgmtNos.forEach((row, index) => {
    const mgmtNo = String(row[0]).trim();
    if (latestProgressMap.has(mgmtNo)) {
      const { progress, editorName, timestamp } = latestProgressMap.get(mgmtNo);
      const rowIndex = index + 2;
      
      progressUpdates.push({row: rowIndex, value: progress});
      editorUpdates.push({row: rowIndex, value: editorName});
      timestampUpdates.push({row: rowIndex, value: timestamp});
    }
  });

  progressUpdates.forEach(update => mainSheet.getRange(update.row, MAIN_SHEET_PROGRESS_COL).setValue(update.value));
  editorUpdates.forEach(update => mainSheet.getRange(update.row, MAIN_SHEET_PROGRESS_EDITOR_COL).setValue(update.value));
  timestampUpdates.forEach(update => mainSheet.getRange(update.row, MAIN_SHEET_UPDATE_TS_COL).setValue(update.value));
  
  if (mainLastRow > 1 && timestampUpdates.length > 0) {
    mainSheet.getRange(2, MAIN_SHEET_UPDATE_TS_COL, mainLastRow - 1, 1).setNumberFormat("yyyy-MM-dd HH:mm");
  }
}

function syncProgressFromMainToInput() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!mainSheet) return;
  const inputSheets = ss.getSheets().filter(s => s.getName().startsWith(INPUT_SHEET_PREFIX));
  const mainLastRow = mainSheet.getLastRow();
  if (mainLastRow <= 1) return;
  const mainData = mainSheet.getRange(2, 1, mainLastRow - 1, Math.max(MAIN_SHEET_PROGRESS_COL, MAIN_SHEET_UPDATE_TS_COL)).getValues();
  const mainProgressMap = new Map();
  mainData.forEach(row => {
    const mgmtNo = String(row[MAIN_SHEET_MGMT_NO_COL - 1]).trim();
    if (mgmtNo) {
      mainProgressMap.set(mgmtNo, {
        progress: row[MAIN_SHEET_PROGRESS_COL - 1] || "未着手",
        timestamp: row[MAIN_SHEET_UPDATE_TS_COL - 1]
      });
    }
  });

  inputSheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    const range = sheet.getRange(3, 1, lastRow - 2, INPUT_SHEET_TIMESTAMP_COL);
    const values = range.getValues();
    let updated = false;
    
    const newValues = values.map(row => {
      const mgmtNo = String(row[INPUT_SHEET_MGMT_NO_COL - 1]).trim();
      if (mainProgressMap.has(mgmtNo)) {
        const { progress, timestamp } = mainProgressMap.get(mgmtNo);
        const mainDate = timestamp ? new Date(timestamp) : null;
        const localDate = row[INPUT_SHEET_TIMESTAMP_COL - 1] ? new Date(row[INPUT_SHEET_TIMESTAMP_COL - 1]) : null;
        
        const shouldUpdate = (progress === '機番重複' && row[INPUT_SHEET_PROGRESS_COL - 1] !== '機番重複') || 
                             (!localDate || (mainDate && mainDate > localDate));

        if (shouldUpdate) {
          if (row[INPUT_SHEET_PROGRESS_COL - 1] !== progress) {
            row[INPUT_SHEET_PROGRESS_COL - 1] = progress;
            updated = true;
          }
        }
      }
      return row;
    });

    if (updated) {
      range.setValues(newValues);
    }
  });
}

function syncProgressToMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  const masterSheet = ss.getSheetByName(KIBAN_MASTER_SHEET_NAME);
  if (!mainSheet || !masterSheet) return;
  const mainLastRow = mainSheet.getLastRow();
  const masterLastRow = masterSheet.getLastRow();
  if (mainLastRow <= 1 || masterLastRow <= 1) return;
  
  const mainData = mainSheet.getRange(2, 1, mainLastRow - 1, MAIN_SHEET_PROGRESS_COL).getValues();
  const progressMap = new Map(mainData.map(row => [String(row[MAIN_SHEET_MGMT_NO_COL - 1]).trim(), row[MAIN_SHEET_PROGRESS_COL - 1] || "未着手"]));
  
  const masterRange = masterSheet.getRange(2, 1, masterLastRow - 1, KIBAN_MASTER_PROGRESS_COL);
  const masterValues = masterRange.getValues();
  let updated = false;

  const newValues = masterValues.map(row => {
    const mgmtNo = String(row[KIBAN_MASTER_MGMT_NO_COL - 1]).trim();
    if (progressMap.has(mgmtNo)) {
      const mainProgress = progressMap.get(mgmtNo);
      if (row[KIBAN_MASTER_PROGRESS_COL - 1] !== mainProgress) {
        row[KIBAN_MASTER_PROGRESS_COL - 1] = mainProgress;
        updated = true;
      }
    }
    return [row[KIBAN_MASTER_PROGRESS_COL - 1]];
  });

  if (updated) {
    masterSheet.getRange(2, KIBAN_MASTER_PROGRESS_COL, newValues.length, 1).setValues(newValues);
  }
}

// =================================================================================
// === 色付け・視覚的補助 (最終レイアウト対応版) ===
// =================================================================================
function colorizeLaborInputColumnsOnHolidaysAndWeekends() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const holidayAndWeekendColor = DUPLICATE_HIGHLIGHT_COLOR;
  const defaultBackgroundColor = DEFAULT_CELL_BACKGROUND_COLOR;
  const calendar = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
  if (!calendar) return;

  const today = new Date();
  const events = calendar.getEvents(new Date(today.getFullYear() - 1, 0, 1), new Date(today.getFullYear() + 1, 11, 31));
  const holidayDates = new Set(events.map(e => formatDateForComparison(e.getStartTime())));

  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(INPUT_SHEET_PREFIX));
  inputSheets.forEach(sheet => {
    const lastCol = sheet.getLastColumn();
    if (lastCol < INPUT_SHEET_LABOR_START_COL) return;

    const headerDates = sheet.getRange(1, INPUT_SHEET_LABOR_START_COL, 1, lastCol - INPUT_SHEET_LABOR_START_COL + 1).getValues()[0];
    const rangeToColor = sheet.getRange(1, INPUT_SHEET_LABOR_START_COL, sheet.getMaxRows(), lastCol - INPUT_SHEET_LABOR_START_COL + 1);
    const backgrounds = rangeToColor.getBackgrounds();
    
    headerDates.forEach((date, i) => {
      if(date){
        const day = new Date(date).getDay();
        const isHoliday = holidayDates.has(formatDateForComparison(new Date(date))) || day === 0 || day === 6;
        if (isHoliday) {
          for (let j = 0; j < backgrounds.length; j++) {
            backgrounds[j][i] = holidayAndWeekendColor;
          }
        }
      }
    });
    rangeToColor.setBackgrounds(backgrounds);
  });
}

function colorizeManagementNoByProgressInMainSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  
  const colorRules = { "未着手": "#ffcfc9", "配置済": "#d4edbc", "ACS済": "#bfe1f6", "日精済": "#ffe5a0", "係り中": "#e6cff2", "機番重複": DUPLICATE_HIGHLIGHT_COLOR, "保留": "#c6dbe1" };
  const range = sheet.getRange(2, 1, lastRow - 1, MAIN_SHEET_PROGRESS_COL);
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();

  const newBgs = values.map((row, i) => {
    const progress = String(row[MAIN_SHEET_PROGRESS_COL - 1]).trim();
    // 機番重複の背景色は checkAndHandleDuplicateMachineNumbers で設定されるため、ここでは上書きしない
    // ただし、重複が解消された場合は、新しい進捗の色を適用する必要がある
    const kiban = String(row[MAIN_SHEET_KIBAN_COL-1]).trim();
    const currentBg = backgrounds[i][MAIN_SHEET_MGMT_NO_COL - 1];

    if (progress === '機番重複') {
      return [DUPLICATE_HIGHLIGHT_COLOR];
    }
    return [colorRules[progress] || DEFAULT_CELL_BACKGROUND_COLOR];
  });
  
  sheet.getRange(2, MAIN_SHEET_MGMT_NO_COL, newBgs.length, 1).setBackgrounds(newBgs);
}

function colorizeProgressColumnInMainSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  
  const colorRules = { "未着手": "#ffcfc9", "配置済": "#d4edbc", "ACS済": "#bfe1f6", "日精済": "#ffe5a0", "係り中": "#e6cff2", "機番重複": DUPLICATE_HIGHLIGHT_COLOR, "保留": "#c6dbe1" };
  const range = sheet.getRange(2, MAIN_SHEET_PROGRESS_COL, lastRow - 1, 1);
  const values = range.getValues();
  
  const newBgs = values.map(row => {
    const progress = String(row[0]).trim();
    return [colorRules[progress] || DEFAULT_CELL_BACKGROUND_COLOR];
  });
  
  range.setBackgrounds(newBgs);
}

function colorizeManagementNoInInputSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(INPUT_SHEET_PREFIX));
  const colorRules = { "未着手": "#ffcfc9", "配置済": "#d4edbc", "ACS済": "#bfe1f6", "日精済": "#ffe5a0", "係り中": "#e6cff2", "機番重複": DUPLICATE_HIGHLIGHT_COLOR, "保留": "#c6dbe1" };

  inputSheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    const range = sheet.getRange(3, 1, lastRow - 2, INPUT_SHEET_PROGRESS_COL);
    const values = range.getValues();
    const backgrounds = range.getBackgrounds();

    const newBgs = values.map((row, i) => {
      const progress = String(row[INPUT_SHEET_PROGRESS_COL - 1]).trim();
      const currentBg = backgrounds[i][0];
      if (progress === '機番重複') {
         return [DUPLICATE_HIGHLIGHT_COLOR];
      }
      return [colorRules[progress] || DEFAULT_CELL_BACKGROUND_COLOR];
    });

    sheet.getRange(3, INPUT_SHEET_MGMT_NO_COL, newBgs.length, 1).setBackgrounds(newBgs);
  });
}

function colorizeTantoushaCellInInputSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(INPUT_SHEET_PREFIX));
  const colorRules = { "志賀": "#ffcfc9", "遠藤": "#d4edbc", "小板橋": "#bfe1f6" };

  inputSheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    const range = sheet.getRange(3, 1, lastRow - 2, INPUT_SHEET_TANTOU_COL);
    const values = range.getValues();
    const backgrounds = range.getBackgrounds();
    
    const newBgs = values.map((row, i) => {
      const tantousha = String(row[INPUT_SHEET_TANTOU_COL - 1]).trim();
      return [colorRules[tantousha] || DEFAULT_CELL_BACKGROUND_COLOR];
    });
    
    sheet.getRange(3, INPUT_SHEET_TANTOU_COL, newBgs.length, 1).setBackgrounds(newBgs);
  });
}

function colorizeToiawaseCellInInputSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(INPUT_SHEET_PREFIX));
  const colorRules = { "問合済": "#ffcfc9", "回答済": "#bfe1f6" };
  
  inputSheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    const range = sheet.getRange(3, 1, lastRow - 2, INPUT_SHEET_TOIAWASE_COL);
    const values = range.getValues();
    
    const newBgs = values.map((row, i) => {
      const toiawase = String(row[INPUT_SHEET_TOIAWASE_COL - 1]).trim();
      return [colorRules[toiawase] || DEFAULT_CELL_BACKGROUND_COLOR];
    });
    
    sheet.getRange(3, INPUT_SHEET_TOIAWASE_COL, newBgs.length, 1).setBackgrounds(newBgs);
  });
}


// =================================================================================
// === 機番重複チェック関連
// =================================================================================
function checkAndHandleDuplicateMachineNumbers() {
  handleDuplicateMachineNumbers(MAIN_SHEET_NAME, MAIN_SHEET_KIBAN_COL, MAIN_SHEET_PROGRESS_COL, '機番重複');
  handleDuplicateMachineNumbersInInputSheets(INPUT_SHEET_KIBAN_COL, INPUT_SHEET_PROGRESS_COL, '機番重複');
}

function processDuplicatesOnSheet(sheetName, machineNoColumnIndex, progressColumnIndex, duplicateText, startRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return;

  const dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, progressColumnIndex);
  const values = dataRange.getValues();

  const seenKibans = new Set();
  let hasChanges = false;

  const newProgressValues = values.map(row => {
    const kiban = String(row[machineNoColumnIndex - 1]).trim();
    let currentProgress = row[progressColumnIndex - 1];

    if (!kiban) {
      if (currentProgress === duplicateText) {
        currentProgress = "未着手";
        hasChanges = true;
      }
    } else {
      if (seenKibans.has(kiban)) {
        if (currentProgress !== duplicateText) {
          currentProgress = duplicateText;
          hasChanges = true;
        }
      } else {
        seenKibans.add(kiban);
        if (currentProgress === duplicateText) {
          currentProgress = "未着手";
          hasChanges = true;
        }
      }
    }
    return [currentProgress];
  });

  if (hasChanges) {
    sheet.getRange(startRow, progressColumnIndex, newProgressValues.length, 1).setValues(newProgressValues);
  }
}

function handleDuplicateMachineNumbers(sheetName, machineNoColumnIndex, progressColumnIndex, duplicateText) {
  processDuplicatesOnSheet(sheetName, machineNoColumnIndex, progressColumnIndex, duplicateText, 2);
}

function handleDuplicateMachineNumbersInInputSheets(machineNoColumnIndex, progressColumnIndex, duplicateText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(INPUT_SHEET_PREFIX));
  inputSheets.forEach(sheet => {
    processDuplicatesOnSheet(sheet.getName(), machineNoColumnIndex, progressColumnIndex, duplicateText, 3);
  });
}


// =================================================================================
// === イベントハンドラ (司令塔) (最終レイアウト対応版) ===
// =================================================================================
function flowManager(e) {
  if (!e || !e.source || !e.range) { return; }
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const sheetName = sheet.getName();
  
  SpreadsheetApp.getActiveSpreadsheet().toast('処理を開始します... (' + sheetName + 'シート編集中)', "情報", 3);
  let actionPerformed = false;

  if (sheetName === KIBAN_MASTER_SHEET_NAME) {
    updateMainSheetFromMaster();
    rebuildInputSheetsFromMainOptimized();
    actionPerformed = true;
  } else if (sheetName === MAIN_SHEET_NAME) {
    // ★★★ ここからが修正箇所です ★★★
    const editedCol = range.getColumn();
    const editedRow = range.getRow();

    // もし進捗列(M列)が編集されたら、同じ行の更新日時(P列)を現在の時刻に更新する
    if (editedCol === MAIN_SHEET_PROGRESS_COL && editedRow >= 2) {
      sheet.getRange(editedRow, MAIN_SHEET_UPDATE_TS_COL).setValue(new Date());
    }
    // ★★★ ここまでが修正箇所です ★★★

    if (editedCol === MAIN_SHEET_TANTOUSHA_COL || editedCol === MAIN_SHEET_TOIAWASE_COL) {
      syncContactInfoFromMainToInputSheets();
    }
    actionPerformed = true;
  } else if (sheetName.startsWith(INPUT_SHEET_PREFIX)) {
    if (range.getColumn() === INPUT_SHEET_PROGRESS_COL && range.getRow() >= 3) {
      sheet.getRange(range.getRow(), INPUT_SHEET_TIMESTAMP_COL).setValue(new Date());
    }
    syncProgressFromInputToMain();
    updateMainSheetLaborTotal();
    actionPerformed = true;
  } else if (sheetName === '生産管理マスタ') {
    syncAssemblyStartDateFromProdMaster();
    actionPerformed = true;
  }

  if (actionPerformed) {
    // 全体的な更新処理をまとめる
    applyProgressPlaceholder();
    batchUpdateCompletionDates();
    syncProgressToMaster();
    markOrphanedRowsInInputSheets();
    
    checkAndHandleDuplicateMachineNumbers();
    syncProgressFromMainToInput();

    colorizeManagementNoByProgressInMainSheet();
    colorizeProgressColumnInMainSheet();
    colorizeManagementNoInInputSheets();
    colorizeTantoushaCellInInputSheets();
    colorizeToiawaseCellInInputSheets();
    SpreadsheetApp.getActiveSpreadsheet().toast("自動処理が完了しました。", "完了", 3);
  }
}

function syncContactInfoFromMainToInputSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!mainSheet) return;
  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(INPUT_SHEET_PREFIX));
  const mainLastRow = mainSheet.getLastRow();
  if (mainLastRow <= 1) return;
  
  const mainData = mainSheet.getRange(2, 1, mainLastRow - 1, Math.max(MAIN_SHEET_TANTOUSHA_COL, MAIN_SHEET_TOIAWASE_COL)).getValues();
  const mainInfoMap = new Map();
  mainData.forEach(row => {
    const mgmtNo = String(row[MAIN_SHEET_MGMT_NO_COL - 1]).trim();
    if (mgmtNo) {
      mainInfoMap.set(mgmtNo, {
        tantousha: row[MAIN_SHEET_TANTOUSHA_COL - 1],
        toiawase: row[MAIN_SHEET_TOIAWASE_COL - 1]
      });
    }
  });

  inputSheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    const range = sheet.getRange(3, 1, lastRow - 2, Math.max(INPUT_SHEET_TANTOU_COL, INPUT_SHEET_TOIAWASE_COL));
    const values = range.getValues();
    let updated = false;

    const newValues = values.map(row => {
      const mgmtNo = String(row[INPUT_SHEET_MGMT_NO_COL - 1]).trim();
      if (mainInfoMap.has(mgmtNo)) {
        const info = mainInfoMap.get(mgmtNo);
        if (row[INPUT_SHEET_TANTOU_COL - 1] !== info.tantousha) {
          row[INPUT_SHEET_TANTOU_COL - 1] = info.tantousha;
          updated = true;
        }
        if (row[INPUT_SHEET_TOIAWASE_COL - 1] !== info.toiawase) {
          row[INPUT_SHEET_TOIAWASE_COL - 1] = info.toiawase;
          updated = true;
        }
      }
      return row; // このままだと全行を更新してしまうので修正
    });

    if (updated) {
       const tantouValues = values.map(row => [row[INPUT_SHEET_TANTOU_COL - 1]]);
       const toiawaseValues = values.map(row => [row[INPUT_SHEET_TOIAWASE_COL - 1]]);
       sheet.getRange(3, INPUT_SHEET_TANTOU_COL, tantouValues.length, 1).setValues(tantouValues);
       sheet.getRange(3, INPUT_SHEET_TOIAWASE_COL, toiawaseValues.length, 1).setValues(toiawaseValues);
    }
  });
}

// =================================================================================
// === カスタムメニューと追加機能
// =================================================================================
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('カスタムメニュー')
      .addItem('操作パネルを開く (工数シート月表示)', 'showControlSidebar')
      .addSeparator()
      .addItem('全機番の資料フォルダ作成 (機番マスタH列)', 'bulkCreateKibanFolders')
      .addItem('全機種シリーズのフォルダ作成 (機番マスタI列)', 'bulkCreateSeriesFolders')
      .addToUi();
}

function showControlSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('操作パネル');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getMonthsFromLaborSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(INPUT_SHEET_PREFIX));
  const months = new Set();
  
  inputSheets.forEach(sheet => {
    const lastCol = sheet.getLastColumn();
    if (lastCol < INPUT_SHEET_LABOR_START_COL) return;
    const headerDates = sheet.getRange(1, INPUT_SHEET_LABOR_START_COL, 1, lastCol - INPUT_SHEET_LABOR_START_COL + 1).getValues()[0];
    headerDates.forEach(date => {
      if (date instanceof Date && !isNaN(date)) {
        months.add(date.getFullYear() + '-' + date.getMonth());
      }
    });
  });
  
  return Array.from(months).map(m => {
    const [year, month] = m.split('-');
    return { text: `<span class="math-inline">\{year\}年</span>{parseInt(month, 10) + 1}月`, value: m };
  }).sort((a, b) => new Date(a.value.split('-')[0], a.value.split('-')[1]) - new Date(b.value.split('-')[0], b.value.split('-')[1]));
}

function filterLaborSheetColumnsByMonth(selectedMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(INPUT_SHEET_PREFIX));
  
  inputSheets.forEach(sheet => {
    const lastCol = sheet.getLastColumn();
    if (lastCol < INPUT_SHEET_LABOR_START_COL) return;
    sheet.showColumns(INPUT_SHEET_LABOR_START_COL, lastCol - INPUT_SHEET_LABOR_START_COL + 1);
    
    if (selectedMonth !== "all") {
      const [targetYear, targetMonth] = selectedMonth.split('-').map(Number);
      const headerDates = sheet.getRange(1, INPUT_SHEET_LABOR_START_COL, 1, lastCol - INPUT_SHEET_LABOR_START_COL + 1).getValues()[0];
      headerDates.forEach((date, i) => {
        if(date instanceof Date && !isNaN(date)){
          const d = new Date(date);
          if (!(d.getFullYear() === targetYear && d.getMonth() === targetMonth)) {
            sheet.hideColumns(INPUT_SHEET_LABOR_START_COL + i);
          }
        }
      });
    }
  });
}

// =================================================================================
// === Google Drive フォルダ作成関連機能
// =================================================================================
function bulkCreateKibanFolders() { bulkCreateFoldersInKibanMasterAndLinkToMain(); }
function bulkCreateSeriesFolders() { bulkCreateSeriesFoldersForKibanMaster(); }

function bulkCreateFoldersInKibanMasterAndLinkToMain() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(KIBAN_MASTER_SHEET_NAME);
  if (!masterSheet) return;
  const lastRow = masterSheet.getLastRow();
  if (lastRow < 2) return;
  const range = masterSheet.getRange(2, 1, lastRow - 1, KIBAN_MASTER_FOLDER_URL_COL);
  const values = range.getValues();
  const formulas = range.getFormulasR1C1();

  values.forEach((row, i) => {
    const kiban = String(row[KIBAN_MASTER_KIBAN_COL - 1]).trim();
    if (kiban && !formulas[i][KIBAN_MASTER_FOLDER_URL_COL -1]) {
      const folderResult = getOrCreateMachineFolder_(kiban, REFERENCE_MATERIAL_PARENT_FOLDER_ID);
      if (folderResult?.folder) {
        formulas[i][KIBAN_MASTER_FOLDER_URL_COL - 1] = `=HYPERLINK("<span class="math-inline">\{folderResult\.folder\.getUrl\(\)\}","</span>{kiban}")`;
      }
    }
  });
  range.setFormulasR1C1(formulas);
  updateMainSheetFromMaster();
}

function bulkCreateSeriesFoldersForKibanMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(KIBAN_MASTER_SHEET_NAME);
  if (!masterSheet) return;
  const lastRow = masterSheet.getLastRow();
  if (lastRow < 2) return;
  const range = masterSheet.getRange(2, 1, lastRow - 1, KIBAN_MASTER_SERIES_FOLDER_URL_COL);
  const values = range.getValues();
  const formulas = range.getFormulasR1C1();

  values.forEach((row, i) => {
    const model = String(row[KIBAN_MASTER_MODEL_COL - 1]).trim();
    const seriesName = extractSeriesPlusInitialNumber_(model);
    if (seriesName && !formulas[i][KIBAN_MASTER_SERIES_FOLDER_URL_COL - 1]) {
      const folderResult = getOrCreateMachineFolder_(seriesName, SERIES_MODEL_PARENT_FOLDER_ID);
      if (folderResult?.folder) {
        formulas[i][KIBAN_MASTER_SERIES_FOLDER_URL_COL - 1] = `=HYPERLINK("<span class="math-inline">\{folderResult\.folder\.getUrl\(\)\}","</span>{seriesName}")`;
      }
    }
  });
  range.setFormulasR1C1(formulas);
  updateMainSheetFromMaster();
}

function getOrCreateMachineFolder_(name, parentFolderId) {
  if (!name || !parentFolderId) return null;
  try {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const folders = parentFolder.getFoldersByName(name);
    if (folders.hasNext()) {
      return { folder: folders.next(), isNew: false };
    }
    return { folder: parentFolder.createFolder(name), isNew: true };
  } catch (e) {
    Logger.log(`フォルダ作成/取得エラー: ${e.message}`);
    return null;
  }
}

function extractSeriesPlusInitialNumber_(modelString, kibanString) {
  if (!modelString || typeof modelString !== 'string') return null;
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
    if (relevantModelString.toUpperCase().startsWith(extractedName.toUpperCase()) && match[1].length >= 2 && match[2].length >= 1) {
      return extractedName;
    }
  }
  return null;
}

// =================================================================================
// === 週次バックアップ機能
// =================================================================================
function createWeeklyBackup() {
  if (!BACKUP_PARENT_FOLDER_ID) return;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const originalName = ss.getName();
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm");
    const backupFileName = `【Backup】<span class="math-inline">\{originalName\}\_</span>{timestamp}`;
    const parentFolder = DriveApp.getFolderById(BACKUP_PARENT_FOLDER_ID);
    DriveApp.getFileById(ss.getId()).makeCopy(backupFileName, parentFolder);

    const files = parentFolder.getFiles();
    const backupFiles = [];
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().startsWith(`【Backup】${originalName}`) && file.getMimeType() === MimeType.GOOGLE_SHEETS) {
        backupFiles.push(file);
      }
    }

    const backupsToKeep = 5;
    if (backupFiles.length > backupsToKeep) {
      backupFiles.sort((a, b) => a.getDateCreated() - b.getDateCreated());
      for (let i = 0; i < backupFiles.length - backupsToKeep; i++) {
        backupFiles[i].setTrashed(true);
      }
    }
  } catch (e) {
    Logger.log(`バックアップエラー: ${e.message}`);
  }
}

// =================================================================================
// === 工数シートへの次月日付列 自動追加機能
// =================================================================================
function addNextMonthDateColumnsToLaborSheets() {
  const today = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(INPUT_SHEET_PREFIX));
  const nextMonthDateObj = new Date(today.getFullYear(), today.getMonth() + 1, 1);
  const nextYear = nextMonthDateObj.getFullYear();
  const nextMonth_0_indexed = nextMonthDateObj.getMonth();

  inputSheets.forEach(sheet => {
    const lastCol = sheet.getLastColumn();
    let nextMonthAlreadyExists = false;
    if (lastCol >= INPUT_SHEET_LABOR_START_COL) {
      const headerValues = sheet.getRange(1, INPUT_SHEET_LABOR_START_COL, 1, lastCol - INPUT_SHEET_LABOR_START_COL + 1).getValues()[0];
      if (headerValues.some(date => date instanceof Date && !isNaN(date) && date.getFullYear() === nextYear && date.getMonth() === nextMonth_0_indexed)) {
        nextMonthAlreadyExists = true;
      }
    }
    
    if (nextMonthAlreadyExists) return;
    
    const daysInNextMonth = new Date(nextYear, nextMonth_0_indexed + 1, 0).getDate();
    const newDateHeaders = Array.from({ length: daysInNextMonth }, (_, i) => new Date(nextYear, nextMonth_0_indexed, i + 1));
    
    if (newDateHeaders.length > 0) {
      const startCol = sheet.getLastColumn() + 1;
      const range = sheet.getRange(1, startCol, 1, newDateHeaders.length);
      range.setValues([newDateHeaders]).setNumberFormat("M/d");
    }
  });
}/**
 * Nissei Spreadsheet Script
 * スプレッドシート操作用のGoogle Apps Script
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Nissei Tools')
    .addItem('データ処理実行', 'processData')
    .addItem('レポート生成', 'generateReport')
    .addToUi();
}

function processData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  console.log('データ処理を開始します...');
  
  // ここにスプレッドシート処理のロジックを追加
  sheet.getRange('A1').setValue('処理完了: ' + new Date());
}

function generateReport() {
  console.log('レポート生成中...');
  
  // レポート生成のロジックを追加
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('B1').setValue('レポート生成: ' + new Date());
}