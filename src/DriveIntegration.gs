/**
 * DriveIntegration.gs
 * Google Driveとの連携機能（フォルダ作成、バックアップ）を管理します。
 */

// =================================================================================
// === Google Drive フォルダ作成関連機能 ===
// =================================================================================

/**
 * 【診断用】メインのフォルダ作成処理に詳細なログを追加したものです。
 */
function bulkCreateMaterialFolders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("========= フォルダ作成処理（診断モード）を開始します =========");
  try {
    const mainSheet = new MainSheet();
    const sheet = mainSheet.getSheet();
    const indices = mainSheet.indices;

    Logger.log(`必須列のインデックス: KIBAN=${indices.KIBAN}, KIBAN_URL=${indices.KIBAN_URL}, MODEL=${indices.MODEL}, SERIES_URL=${indices.SERIES_URL}`);
    if (!indices.KIBAN || !indices.KIBAN_URL || !indices.MODEL || !indices.SERIES_URL) {
      throw new Error(`必要な列が見つかりません。`);
    }

    const lastRow = mainSheet.getLastRow();
    Logger.log(`最終行: ${lastRow}, データ開始行: ${mainSheet.startRow}`);
    if (lastRow < mainSheet.startRow) {
       Logger.log("データ行がないため処理を終了します。");
       return;
    }

    const range = sheet.getRange(mainSheet.startRow, 1, lastRow - mainSheet.startRow + 1, mainSheet.getLastColumn());
    const values = range.getValues();
    const formulas = range.getFormulas();
    Logger.log(`${values.length}行のデータを読み込みました。`);

    const processedItems = new Set();
    const folderCreationTasks = [
      { type: 'KIBAN', valueCol: indices.KIBAN, linkCol: indices.KIBAN_URL, parentFolderId: CONFIG.FOLDERS.REFERENCE_MATERIAL_PARENT },
      { type: 'MODEL', valueCol: indices.MODEL, linkCol: indices.SERIES_URL, parentFolderId: CONFIG.FOLDERS.SERIES_MODEL_PARENT }
    ];

    folderCreationTasks.forEach(task => {
      Logger.log(`--- [${task.type}] のタスクを開始します ---`);
      processedItems.clear();
      values.forEach((row, i) => {
        const currentRowNum = mainSheet.startRow + i;
        const originalValue = String(row[task.valueCol - 1]).trim();
        if (!originalValue) return;

        let groupValue, folderName;
        if (task.type === 'MODEL') {
          const match = originalValue.match(/^[A-Za-z]+[0-9]+/);
          groupValue = match ? match[0] : originalValue;
          folderName = groupValue;
        } else {
          groupValue = originalValue;
          folderName = originalValue;
        }

        if (processedItems.has(groupValue)) return; // 既に処理済みの場合はスキップ
        
        Logger.log(`${currentRowNum}行目: 値「${originalValue}」を検出。グループ「${groupValue}」として処理します。`);
        processedItems.add(groupValue); // 重複処理を防ぐためにセットに追加

        const folderResult = getOrCreateFolder_(folderName, task.parentFolderId);
        if (folderResult && folderResult.folder) {
          const url = folderResult.folder.getUrl();
          Logger.log(`  -> フォルダを${folderResult.isNew ? '新規作成' : '取得'}しました。URL: ${url}`);
          
          if (task.type === 'MODEL') {
            updateLinksForModelPrefix_(values, formulas, groupValue, task.valueCol, task.linkCol, url);
          } else {
            updateLinksForSameValue_(values, formulas, groupValue, task.valueCol, task.linkCol, url);
          }
        } else {
          Logger.log(`  -> フォルダの作成/取得に失敗しました。`);
        }
      });
    });

    Logger.log("全タスク完了。シートにデータを書き込みます...");
    const outputData = values.map((row, i) => row.map((cell, j) => formulas[i][j] || cell));
    range.setValues(outputData);

    ss.toast("資料フォルダの作成とリンク設定が完了しました。");
    Logger.log("========= 処理が正常に完了しました =========");

  } catch (error) {
    Logger.log(`エラー発生: ${error.stack}`);
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
 * (新規追加)
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