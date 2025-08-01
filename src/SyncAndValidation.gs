// =================================================================================
// === 進捗同期と重複チェック関連 ===
// =================================================================================
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