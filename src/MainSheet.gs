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