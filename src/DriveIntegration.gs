/**
 * DriveIntegration.gs
 * Google Driveとの連携機能（フォルダ作成、バックアップ）を管理します。
 * (管理表の作成タイミングを分離)
 */

// =================================================================================
// === Google Drive フォルダ作成関連機能 ===
// =================================================================================

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
          uniqueModels.add(match ? match[0] : model);
        }
      }
    });

    const kibanUrlMap = new Map();
    uniqueKibans.forEach(kiban => {
      const folderResult = getOrCreateFolder_(kiban, CONFIG.FOLDERS.REFERENCE_MATERIAL_PARENT);
      if (folderResult && folderResult.folder) {
        kibanUrlMap.set(kiban, folderResult.folder.getUrl());
      }
    });

    const modelUrlMap = new Map();
    uniqueModels.forEach(modelPrefix => {
      const folderResult = getOrCreateFolder_(modelPrefix, CONFIG.FOLDERS.SERIES_MODEL_PARENT);
      if (folderResult && folderResult.folder) {
        modelUrlMap.set(modelPrefix, folderResult.folder.getUrl());
      }
    });
    
    const outputData = values.map((row, i) => row.map((cell, j) => formulas[i][j] || cell));
    let modified = false;
    values.forEach((row, i) => {
      const kiban = String(row[indices.KIBAN - 1]).trim();
      if (kibanUrlMap.has(kiban) && !formulas[i][indices.KIBAN_URL - 1]) {
        outputData[i][indices.KIBAN_URL - 1] = createHyperlinkFormula(kibanUrlMap.get(kiban), kiban);
        modified = true;
      }
      const model = String(row[indices.MODEL - 1]).trim();
      if (model) {
        const match = model.match(/^[A-Za-z]+[0-9]+/);
        const groupValue = match ? match[0] : model;
        if (modelUrlMap.has(groupValue) && !formulas[i][indices.SERIES_URL - 1]) {
          outputData[i][indices.SERIES_URL - 1] = createHyperlinkFormula(modelUrlMap.get(groupValue), groupValue);
          modified = true;
        }
      }
    });

    if (modified) {
      range.setValues(outputData);
      ss.toast("資料フォルダの作成とリンク設定が完了しました。");
    } else {
      ss.toast("すべてのリンクは既に設定済みです。");
    }

  } catch (error) {
    Logger.log(error.stack);
    ss.toast(`エラー: ${error.message}`);
  }
}

/**
 * テンプレートのスプレッドシートをコピーし、情報を書き込んでフォルダに保存する
 * (エラー発生時に最大4回まで自動で再試行する機能を追加)
 */
function createManagementSheet_(targetFolder, kiban, model) {
  const templateId = CONFIG.TEMPLATES.MANAGEMENT_SHEET_TEMPLATE_ID;
  if (!templateId || templateId.includes("...")) {
    Logger.log("テンプレートIDがConfig.gsに設定されていません。");
    return;
  }

  try {
    const templateFile = DriveApp.getFileById(templateId);
    const newFileName = `${kiban}盤配指示図出図管理表`;
    
    // 重複作成を防止
    const existingFiles = targetFolder.getFilesByName(newFileName);
    if (existingFiles.hasNext()) {
      Logger.log(`管理表「${newFileName}」は既に存在するため、作成をスキップします。`);
      return;
    }

    const newFile = templateFile.makeCopy(newFileName, targetFolder);
    const newFileId = newFile.getId();

    let success = false;
    let attempts = 4;
    let waitTime = 2000;

    Utilities.sleep(2000);

    for (let i = 0; i < attempts; i++) {
      try {
        const newSpreadsheet = SpreadsheetApp.openById(newFileId);
        const sheet = newSpreadsheet.getSheets()[0];
        
        sheet.getRange("B4").setValue("機種：" + model);
        sheet.getRange("B5").setValue("製番：" + kiban);
        sheet.getDataRange().setFontFamily("Arial").setFontSize(11);
        SpreadsheetApp.flush();
        
        success = true;
        Logger.log(`管理表「${newFileName}」をフォルダ「${targetFolder.getName()}」に作成し、編集しました。`);
        SpreadsheetApp.getActiveSpreadsheet().toast(`管理表「${newFileName}」を作成しました。`);
        break;
      } catch (e) {
        if (e.message.includes("サービスに接続できなくなりました")) {
          Logger.log(`試行 ${i + 1}/${attempts}: スプレッドシートサービスへの接続に失敗。${waitTime / 1000}秒後に再試行します。`);
          Utilities.sleep(waitTime);
          waitTime *= 2;
        } else {
          throw e;
        }
      }
    }

    if (!success) {
      Logger.log(`${attempts}回の再試行後も管理表の編集に失敗しました。`);
      SpreadsheetApp.getActiveSpreadsheet().toast("管理表の編集に失敗しました。時間をおいて再度お試しください。");
    }

  } catch (e) {
    Logger.log(`管理表の作成または編集中にエラーが発生しました: ${e.message}`);
    SpreadsheetApp.getActiveSpreadsheet().toast("管理表の作成中にエラーが発生しました。ログを確認してください。");
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
    if (!CONFIG.FOLDERS.BACKUP_PARENT) {
      throw new Error("バックアップ用フォルダIDが設定されていません。");
    }
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