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
      return row; // このままだと全行を更新してしまうので修正
    });

    if (updated) {
      range.setValues(newValues);
    }
  });
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
}