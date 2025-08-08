/**
 * DriveIntegration.gs
 * Google Driveとの連携機能（フォルダ作成、バックアップ）を管理します。
 * (効率化・安定化・詳細ログ出力 改訂版)
 */

// =================================================================================
// === Google Drive フォルダ作成関連機能 ===
// =================================================================================

function bulkCreateMaterialFolders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('フォルダ作成処理を開始します。');
  try {
    const mainSheet = new MainSheet();
    const sheet = mainSheet.getSheet();
    const indices = mainSheet.indices;
    Logger.log('メインシートの情報を取得しました。');

    // 必須列の存在チェック
    const requiredColumns = {
      KIBAN: indices.KIBAN, KIBAN_URL: indices.KIBAN_URL,
      MODEL: indices.MODEL, SERIES_URL: indices.SERIES_URL
    };
    for (const [key, value] of Object.entries(requiredColumns)) {
      if (!value) throw new Error(`必須列「${key}」が見つかりません。`);
    }

    const lastRow = mainSheet.getLastRow();
    if (lastRow < mainSheet.startRow) {
      Logger.log('データが存在しないため処理を終了します。');
      ss.toast('データがありません。');
      return;
    }

    const range = sheet.getRange(mainSheet.startRow, 1, lastRow - mainSheet.startRow + 1, mainSheet.getLastColumn());
    const values = range.getValues();
    const formulas = range.getFormulas();
    Logger.log(`${values.length}行のデータを読み込みました。`);

    // 手順1: リンクが未作成のユニークな「機番」と「機種プレフィックス」を収集
    const uniqueKibans = new Set();
    const uniqueModels = new Set();
    values.forEach((row, i) => {
      if (!formulas[i][indices.KIBAN_URL - 1]) {
        const kiban = String(row[indices.KIBAN - 1]).trim();
        if (kiban) uniqueKibans.add(kiban);
      }
      if (!formulas[i][indices.SERIES_URL - 1]) {
        const model = String(row[indices.MODEL - 1]).trim();
        if (model) {
          const match = model.match(/^[A-Za-z]+[0-9]+/);
          const groupValue = match ? match[0] : model;
          uniqueModels.add(groupValue);
        }
      }
    });
    Logger.log(`リンク作成が必要なユニーク機番: ${[...uniqueKibans].join(', ')}`);
    Logger.log(`リンク作成が必要なユニーク機種: ${[...uniqueModels].join(', ')}`);

    if (uniqueKibans.size === 0 && uniqueModels.size === 0) {
        Logger.log('すべてのリンクは既に設定済みと判断しました。処理を終了します。');
        ss.toast("すべてのリンクは既に設定済みです。");
        return;
    }

    // 手順2: フォルダを一括作成し、URLをマップに保存
    const kibanUrlMap = new Map();
    uniqueKibans.forEach(kiban => {
      const folderResult = getOrCreateFolder_(kiban, CONFIG.FOLDERS.REFERENCE_MATERIAL_PARENT);
      if (folderResult && folderResult.folder) kibanUrlMap.set(kiban, folderResult.folder.getUrl());
    });
    Logger.log(`${kibanUrlMap.size}個の機番フォルダのURLを取得/作成しました。`);

    const modelUrlMap = new Map();
    uniqueModels.forEach(modelPrefix => {
      const folderResult = getOrCreateFolder_(modelPrefix, CONFIG.FOLDERS.SERIES_MODEL_PARENT);
      if (folderResult && folderResult.folder) modelUrlMap.set(modelPrefix, folderResult.folder.getUrl());
    });
    Logger.log(`${modelUrlMap.size}個の機種フォルダのURLを取得/作成しました。`);

    // 手順3: 収集したURLを元に、シートに書き込むデータ（formulas配列）を更新
    let modified = false;
    let linksCreated = 0;
    values.forEach((row, i) => {
      const kiban = String(row[indices.KIBAN - 1]).trim();
      if (kibanUrlMap.has(kiban) && !formulas[i][indices.KIBAN_URL - 1]) {
        formulas[i][indices.KIBAN_URL - 1] = createHyperlinkFormula(kibanUrlMap.get(kiban), kiban);
        modified = true;
        linksCreated++;
      }
      const model = String(row[indices.MODEL - 1]).trim();
      if (model) {
        const match = model.match(/^[A-Za-z]+[0-9]+/);
        const groupValue = match ? match[0] : model;
        if (modelUrlMap.has(groupValue) && !formulas[i][indices.SERIES_URL - 1]) {
          formulas[i][indices.SERIES_URL - 1] = createHyperlinkFormula(modelUrlMap.get(groupValue), groupValue);
          modified = true;
          linksCreated++;
        }
      }
    });
    Logger.log(`${linksCreated}個のリンクを生成しました。`);

    // 手順4: 変更があった場合のみ、更新後のformulas配列をシートに書き戻す
    if (modified) {
      Logger.log('変更があったため、シートに書き込みます...');
      range.setFormulas(formulas);
      Logger.log('書き込みが完了しました。');
      ss.toast("資料フォルダの作成とリンク設定が完了しました。");
    } else {
      Logger.log('シートに書き込む変更はありませんでした。');
      ss.toast("リンクを更新する行はありませんでした。");
    }

  } catch (error) {
    Logger.log(`エラーが発生しました: ${error.stack}`);
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
    Logger.log(`フォルダ作成/取得エラー: ${e.message}`);
    return null;
  }
}

// =================================================================================
// === 週次バックアップ機能 (変更なし) ===
// =================================================================================
function createWeeklyBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    if (!CONFIG.FOLDERS.BACKUP_PARENT) throw new Error("バックアップ用フォルダIDが設定されていません。");
    const originalName = ss.getName();
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMATS.BACKUP_TIMESTAMP);
    const backupFileName = `${CONFIG.BACKUP.PREFIX}${originalName}_${timestamp}`;
    const parentFolder = DriveApp.getFolderById(CONFIG.FOLDERS.BACKUP_PARENT);
    DriveApp.getFileById(ss.getId()).makeCopy(backupFileName, parentFolder);
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