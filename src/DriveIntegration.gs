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
}/**
 * DriveIntegration.gs
 * Google Driveとの連携機能（フォルダ作成、バックアップ）を管理
 */

// =================================================================================
// === Google Drive フォルダ作成関連機能 ===
// =================================================================================

/**
 * メインシートのデータに基づき、機番ごとの資料フォルダを一括作成し、シートにリンクを挿入します。
 */
function bulkCreateKibanFolders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const mainSheet = new MainSheet(); // MainSheetオブジェクトとしてシートを取得
    const sheet = mainSheet.getSheet();
    const indices = mainSheet.indices;

    if (!indices.KIBAN || !indices.KIBAN_URL) {
      throw new Error("「機番」または「機番(リンク)」列が見つかりません。");
    }

    const lastRow = mainSheet.getLastRow();
    if (lastRow < 2) return;

    const range = sheet.getRange(2, 1, lastRow - 1, mainSheet.getLastColumn());
    const values = range.getValues();
    const formulas = range.getFormulasR1C1();

    values.forEach((row, i) => {
      const kiban = safeTrim(row[indices.KIBAN - 1]);
      if (kiban && !formulas[i][indices.KIBAN_URL - 1]) {
        const folderResult = getOrCreateFolder_(kiban, CONFIG.FOLDERS.REFERENCE_MATERIAL_PARENT);
        if (folderResult && folderResult.folder) {
          formulas[i][indices.KIBAN_URL - 1] = createHyperlinkFormula(folderResult.folder.getUrl(), kiban);
        }
      }
    });

    range.setFormulasR1C1(formulas);
    ss.toast("機番フォルダの作成とリンク設定が完了しました。");

  } catch (error) {
    Logger.log(error.stack);
    ss.toast(`エラー: ${error.message}`);
  }
}

/**
 * メインシートのデータに基づき、機種シリーズごとの資料フォルダを一括作成し、シートにリンクを挿入します。
 */
function bulkCreateSeriesFolders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const mainSheet = new MainSheet(); // MainSheetオブジェクトとしてシートを取得
    const sheet = mainSheet.getSheet();
    const indices = mainSheet.indices;
    
    if (!indices.MODEL || !indices.SERIES_URL) {
      throw new Error("「機種」または「STD資料(リンク)」列が見つかりません。");
    }

    const lastRow = mainSheet.getLastRow();
    if (lastRow < 2) return;

    const range = sheet.getRange(2, 1, lastRow - 1, mainSheet.getLastColumn());
    const values = range.getValues();
    const formulas = range.getFormulasR1C1();

    values.forEach((row, i) => {
      const model = safeTrim(row[indices.MODEL - 1]);
      const seriesName = extractSeriesPlusInitialNumber(model);
      if (seriesName && !formulas[i][indices.SERIES_URL - 1]) {
        const folderResult = getOrCreateFolder_(seriesName, CONFIG.FOLDERS.SERIES_MODEL_PARENT);
        if (folderResult && folderResult.folder) {
          formulas[i][indices.SERIES_URL - 1] = createHyperlinkFormula(folderResult.folder.getUrl(), seriesName);
        }
      }
    });

    range.setFormulasR1C1(formulas);
    ss.toast("機種シリーズフォルダの作成とリンク設定が完了しました。");

  } catch (error) {
    Logger.log(error.stack);
    ss.toast(`エラー: ${error.message}`);
  }
}

/**
 * 指定された名前のフォルダを、指定された親フォルダ内に作成または取得する内部関数。
 */
function getOrCreateFolder_(name, parentFolderId) {
  if (!name || !parentFolderId) return null;
  try {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const folders = parentFolder.getFoldersByName(name);
    if (folders.hasNext()) {
      return { folder: folders.next(), isNew: false };
    }
    return { folder: parentFolder.createFolder(name), isNew: true };
  } catch (e) {
    Logger.log(ERROR_MESSAGES.FOLDER_CREATE_ERROR + e.message);
    return null;
  }
}

// =================================================================================
// === 週次バックアップ機能 ===
// =================================================================================
function createWeeklyBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    if (!CONFIG.FOLDERS.BACKUP_PARENT) {
      throw new Error("バックアップ用フォルダIDが設定されていません。");
    }
    
    const originalName = ss.getName();
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMATS.BACKUP_TIMESTAMP);
    const backupFileName = `${CONFIG.BACKUP.PREFIX}${originalName}_${timestamp}`;
    const parentFolder = DriveApp.getFolderById(CONFIG.FOLDERS.BACKUP_PARENT);
    
    DriveApp.getFileById(ss.getId()).makeCopy(backupFileName, parentFolder);

    // 古いバックアップの削除
    const files = parentFolder.getFiles();
    const backupFiles = [];
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().startsWith(`${CONFIG.BACKUP.PREFIX}${originalName}`) && file.getMimeType() === MimeType.GOOGLE_SHEETS) {
        backupFiles.push(file);
      }
    }

    if (backupFiles.length > CONFIG.BACKUP.KEEP_COUNT) {
      backupFiles.sort((a, b) => a.getDateCreated() - b.getDateCreated());
      const toDeleteCount = backupFiles.length - CONFIG.BACKUP.KEEP_COUNT;
      for (let i = 0; i < toDeleteCount; i++) {
        backupFiles[i].setTrashed(true);
      }
    }
    ss.toast("バックアップを作成しました。");
  } catch (e) {
    Logger.log(ERROR_MESSAGES.BACKUP_ERROR + e.message);
    ss.toast(`バックアップエラー: ${e.message}`);
  }
}