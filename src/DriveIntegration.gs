/**
 * DriveIntegration.gs
 * Google Driveとの連携機能（フォルダ作成、バックアップ）を管理します。
 */

// =================================================================================
// === Google Drive フォルダ作成関連機能 ===
// =================================================================================

/**
 * メインシートのデータに基づき、資料フォルダを一括作成し、シートにリンクを挿入します。
 */
function bulkCreateMaterialFolders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const mainSheet = new MainSheet();
    const sheet = mainSheet.getSheet();
    const indices = mainSheet.indices;

    const requiredColumns = {
      KIBAN: indices.KIBAN, KIBAN_URL: indices.KIBAN_URL,
      MODEL: indices.MODEL, SERIES_URL: indices.SERIES_URL
    };
    for (const [key, value] of Object.entries(requiredColumns)) {
      if (!value) throw new Error(`必要な列「${key}」が見つかりません。`);
    }

    const lastRow = mainSheet.getLastRow();
    if (lastRow < mainSheet.startRow) return;

    const range = sheet.getRange(mainSheet.startRow, 1, lastRow - mainSheet.startRow + 1, mainSheet.getLastColumn());
    const values = range.getValues();
    const formulas = range.getFormulas();

    const processedItems = new Set();

    const folderCreationTasks = [
      {
        valueCol: indices.KIBAN,
        linkCol: indices.KIBAN_URL,
        parentFolderId: CONFIG.FOLDERS.REFERENCE_MATERIAL_PARENT
      },
      {
        valueCol: indices.MODEL,
        linkCol: indices.SERIES_URL,
        parentFolderId: CONFIG.FOLDERS.SERIES_MODEL_PARENT
      }
    ];
    
    folderCreationTasks.forEach(task => {
      processedItems.clear();
      values.forEach((row, i) => {
        const value = String(row[task.valueCol - 1]).trim();
        if (value && !processedItems.has(value)) {
          const folderResult = getOrCreateFolder_(value, task.parentFolderId);
          if (folderResult && folderResult.folder) {
            updateLinksForSameValue_(values, formulas, value, task.valueCol, task.linkCol, folderResult.folder.getUrl());
          }
          processedItems.add(value);
        }
      });
    });

    const outputData = values.map((row, i) => row.map((cell, j) => formulas[i][j] || cell));
    range.setValues(outputData);

    ss.toast("資料フォルダの作成とリンク設定が完了しました。");
  } catch (error) {
    Logger.log(error.stack);
    ss.toast(`エラー: ${error.message}`);
  }
}

/**
 * (リファクタリング)
 * 指定された列の値が一致するすべての行に、ハイパーリンクを設定します。
 */
function updateLinksForSameValue_(allValues, allFormulas, valueToMatch, valueColumn, linkColumn, url) {
  allValues.forEach((row, i) => {
    if (String(row[valueColumn - 1]).trim() === valueToMatch) {
      if (!allFormulas[i][linkColumn - 1]) {
        allFormulas[i][linkColumn - 1] = createHyperlinkFormula(url, valueToMatch);
      }
    }
  });
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
    Logger.log(`フォルダ作成/取得エラー: ${e.message}`);
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
    Logger.log(`バックアップエラー: ${e.message}`);
    ss.toast(`バックアップエラー: ${e.message}`);
  }
}