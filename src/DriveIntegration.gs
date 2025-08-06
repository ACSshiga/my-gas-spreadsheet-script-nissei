/**
 * DriveIntegration.gs
 * Google Driveとの連携機能（フォルダ作成、バックアップ）を管理します。
 */

// =================================================================================
// === Google Drive フォルダ作成関連機能 ===
// =================================================================================

/**
 * メインシートのデータに基づき、資料フォルダを一括作成し、シートにリンクを挿入します。
 * 機種(MODEL)については、指定された命名規則（先頭の英語+数字）に基づいてグルーピングされます。
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
        type: 'KIBAN', // 機番は完全一致
        valueCol: indices.KIBAN,
        linkCol: indices.KIBAN_URL,
        parentFolderId: CONFIG.FOLDERS.REFERENCE_MATERIAL_PARENT
      },
      {
        type: 'MODEL', // 機種は前方一致
        valueCol: indices.MODEL,
        linkCol: indices.SERIES_URL,
        parentFolderId: CONFIG.FOLDERS.SERIES_MODEL_PARENT
      }
    ];

    folderCreationTasks.forEach(task => {
      processedItems.clear();
      values.forEach((row, i) => {
        const originalValue = String(row[task.valueCol - 1]).trim();
        if (!originalValue) return;

        let groupValue;
        let folderName;

        if (task.type === 'MODEL') {
          const match = originalValue.match(/^[A-Za-z]+[0-9]+/);
          groupValue = match ? match[0] : originalValue;
          folderName = groupValue;
        } else {
          groupValue = originalValue;
          folderName = originalValue;
        }
        
        if (groupValue && !processedItems.has(groupValue)) {
          const folderResult = getOrCreateFolder_(folderName, task.parentFolderId);
          if (folderResult && folderResult.folder) {
            const url = folderResult.folder.getUrl();
            if (task.type === 'MODEL') {
              updateLinksForModelPrefix_(values, formulas, groupValue, task.valueCol, task.linkCol, url);
            } else {
              updateLinksForSameValue_(values, formulas, groupValue, task.valueCol, task.linkCol, url);
            }
          }
          processedItems.add(groupValue);
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
 * 指定されたプレフィックスに前方一致するすべての行に、ハイパーリンクを設定します。
 */
function updateLinksForModelPrefix_(allValues, allFormulas, prefixToMatch, valueColumn, linkColumn, url) {
  allValues.forEach((row, i) => {
    const fullValue = String(row[valueColumn - 1]).trim();
    if (fullValue.startsWith(prefixToMatch)) {
      if (!allFormulas[i][linkColumn - 1]) {
        allFormulas[i][linkColumn - 1] = createHyperlinkFormula(url, fullValue);
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