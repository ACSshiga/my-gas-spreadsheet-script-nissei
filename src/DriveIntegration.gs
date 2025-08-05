/**
 * DriveIntegration.gs
 * Google Driveとの連携機能（フォルダ作成、バックアップ）を管理します。
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
    const mainSheet = new MainSheet();
    const sheet = mainSheet.getSheet();
    const indices = mainSheet.indices;
    if (!indices.KIBAN || !indices.KIBAN_URL) {
      throw new Error("「機番」または「機番(リンク)」列が見つかりません。");
    }

    const lastRow = mainSheet.getLastRow();
    if (lastRow < mainSheet.startRow) return;
    const range = sheet.getRange(mainSheet.startRow, 1, lastRow - mainSheet.startRow + 1, mainSheet.getLastColumn());
    const values = range.getValues();
    const formulas = range.getFormulasR1C1();

    // 重複する機番でフォルダが複数作成されるのを防ぐため、処理済みの機番を記録
    const processedKiban = new Set();
    values.forEach((row, i) => {
      const kiban = String(row[indices.KIBAN - 1]).trim();
      
      // 機番があり、まだ処理されておらず、リンクが空の場合のみ実行
      if (kiban && !processedKiban.has(kiban)) {
        const folderResult = getOrCreateFolder_(kiban, CONFIG.FOLDERS.REFERENCE_MATERIAL_PARENT);
        if (folderResult && folderResult.folder) {
          // 同じ機番を持つすべての行のリンクを更新
          updateLinksForSameKiban_(values, formulas, kiban, indices, folderResult.folder.getUrl());
        }
        processedKiban.add(kiban);
      }
    });

    // ★★★ 修正箇所 ★★★
    // E列（機番(リンク)列）の数式のみを更新し、他の列に影響を与えないようにします。
    const kibanUrlColumn = indices.KIBAN_URL;
    const formulasForKibanUrl = values.map((row, i) => [formulas[i][kibanUrlColumn - 1]]);
    
    sheet.getRange(mainSheet.startRow, kibanUrlColumn, formulasForKibanUrl.length, 1).setFormulasR1C1(formulasForKibanUrl);
    // ★★★ 修正ここまで ★★★

    ss.toast("機番フォルダの作成とリンク設定が完了しました。");
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