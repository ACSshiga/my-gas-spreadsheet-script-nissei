/**
 * DriveIntegration.gs
 * Google Driveとの連携機能（フォルダ作成、バックアップ）を管理します。
 */

// =================================================================================
// === Google Drive フォルダ作成関連機能 ===
// =================================================================================

/**
 * メインシートのデータに基づき、機番ごと、および機種ごとの資料フォルダを一括作成し、シートにリンクを挿入します。
 */
function bulkCreateMaterialFolders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const mainSheet = new MainSheet();
    const sheet = mainSheet.getSheet();
    const indices = mainSheet.indices;

    // 機能に必要な列が存在するかチェック
    if (!indices.KIBAN || !indices.KIBAN_URL || !indices.MODEL || !indices.SERIES_URL) {
      throw new Error("「機番」「機番(リンク)」「機種」「STD資料(リンク)」のいずれかの列が見つかりません。");
    }

    const lastRow = mainSheet.getLastRow();
    if (lastRow < mainSheet.startRow) return;

    const range = sheet.getRange(mainSheet.startRow, 1, lastRow - mainSheet.startRow + 1, mainSheet.getLastColumn());
    const values = range.getValues();
    const formulas = range.getFormulas(); // R1C1からA1表記に変更

    // 重複処理を防ぐため、処理済みの機番と機種を記録
    const processedKiban = new Set();
    const processedModel = new Set();

    values.forEach((row, i) => {
      const kiban = String(row[indices.KIBAN - 1]).trim();
      const model = String(row[indices.MODEL - 1]).trim();
      
      // 【機番フォルダ作成ロジック】
      // 機番があり、まだ処理されておらず、リンクが空の場合のみ実行
      if (kiban && !processedKiban.has(kiban)) {
        const folderResult = getOrCreateFolder_(kiban, CONFIG.FOLDERS.REFERENCE_MATERIAL_PARENT);
        if (folderResult && folderResult.folder) {
          updateLinksForSameKiban_(values, formulas, kiban, indices, folderResult.folder.getUrl());
        }
        processedKiban.add(kiban);
      }

      // 【STD資料（機種）フォルダ作成ロジック】
      // 機種があり、まだ処理されておらず、リンクが空の場合のみ実行
      if (model && !processedModel.has(model)) {
        const folderResult = getOrCreateFolder_(model, CONFIG.FOLDERS.SERIES_MODEL_PARENT);
        if (folderResult && folderResult.folder) {
          updateLinksForSameModel_(values, formulas, model, indices, folderResult.folder.getUrl());
        }
        processedModel.add(model);
      }
    });

    // 数式と値をマージして、安全にシートへ書き戻す
    const outputData = values.map((row, i) => {
      return row.map((cell, j) => {
        // 更新された数式があればそれを使い、なければ元の値を使う
        return formulas[i][j] || values[i][j];
      });
    });
    range.setValues(outputData);

    ss.toast("資料フォルダの作成とリンク設定が完了しました。");
  } catch (error) {
    Logger.log(error.stack);
    ss.toast(`エラー: ${error.message}`);
  }
}

/**
 * 同じ機番を持つすべての行に、指定されたURLのハイパーリンクを設定します。
 */
function updateLinksForSameKiban_(allValues, allFormulas, kiban, indices, url) {
  allValues.forEach((row, i) => {
    if (String(row[indices.KIBAN - 1]).trim() === kiban) {
      // リンクがまだ設定されていない場合のみ設定
      if (!allFormulas[i][indices.KIBAN_URL - 1]) {
         allFormulas[i][indices.KIBAN_URL - 1] = createHyperlinkFormula(url, kiban);
      }
    }
  });
}

/**
 * ★★★ 新規追加 ★★★
 * 同じ機種を持つすべての行に、指定されたURLのハイパーリンクを設定します。
 */
function updateLinksForSameModel_(allValues, allFormulas, model, indices, url) {
  allValues.forEach((row, i) => {
    if (String(row[indices.MODEL - 1]).trim() === model) {
      // リンクがまだ設定されていない場合のみ設定
      if (!allFormulas[i][indices.SERIES_URL - 1]) {
         allFormulas[i][indices.SERIES_URL - 1] = createHyperlinkFormula(url, model);
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