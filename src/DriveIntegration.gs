// =================================================================================
// === Google Drive フォルダ作成関連機能 ===
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
        formulas[i][KIBAN_MASTER_FOLDER_URL_COL - 1] = `=HYPERLINK("${folderResult.folder.getUrl()}","${kiban}")`;
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
        formulas[i][KIBAN_MASTER_SERIES_FOLDER_URL_COL - 1] = `=HYPERLINK("${folderResult.folder.getUrl()}","${seriesName}")`;
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
// === 週次バックアップ機能 ===
// =================================================================================
function createWeeklyBackup() {
  if (!BACKUP_PARENT_FOLDER_ID) return;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const originalName = ss.getName();
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm");
    const backupFileName = `【Backup】${originalName}_${timestamp}`;
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