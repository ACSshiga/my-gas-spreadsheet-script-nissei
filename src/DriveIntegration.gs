/**
 * DriveIntegration.gs
 * Google Driveとの連携機能（フォルダ作成、バックアップ）を管理します。
 * (顧客用管理シートの自動生成、書式設定機能を追加)
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

    const uniqueKibans = new Map();
    const uniqueModels = new Set();
    values.forEach((row, i) => {
      const kiban = String(row[indices.KIBAN - 1]).trim();
      const model = String(row[indices.MODEL - 1]).trim();
      if (kiban && !formulas[i][indices.KIBAN_URL - 1] && !uniqueKibans.has(kiban)) {
        uniqueKibans.set(kiban, model);
      }
      if (model && !formulas[i][indices.SERIES_URL - 1]) {
        const match = model.match(/^[A-Za-z]+[0-9]+/);
        if(match) uniqueModels.add(match[0]);
      }
    });

    const kibanUrlMap = new Map();
    uniqueKibans.forEach((model, kiban) => {
      const folderResult = getOrCreateFolder_(kiban, CONFIG.FOLDERS.REFERENCE_MATERIAL_PARENT);
      if (folderResult && folderResult.folder) {
        kibanUrlMap.set(kiban, folderResult.folder.getUrl());
        if (folderResult.isNew) {
          createManagementSheet_(folderResult.folder, kiban, model);
        }
      }
    });

    const modelUrlMap = new Map();
    uniqueModels.forEach(modelPrefix => {
      const folderResult = getOrCreateFolder_(modelPrefix, CONFIG.FOLDERS.SERIES_MODEL_PARENT);
      if (folderResult && folderResult.folder) {
        modelUrlMap.set(modelPrefix, folderResult.folder.getUrl());
      }
    });

    let modified = false;
    values.forEach((row, i) => {
      const kiban = String(row[indices.KIBAN - 1]).trim();
      if (kibanUrlMap.has(kiban) && !formulas[i][indices.KIBAN_URL - 1]) {
        formulas[i][indices.KIBAN_URL - 1] = createHyperlinkFormula(kibanUrlMap.get(kiban), kiban);
        modified = true;
      }
      const model = String(row[indices.MODEL - 1]).trim();
      if (model) {
        const match = model.match(/^[A-Za-z]+[0-9]+/);
        const groupValue = match ? match[0] : model;
        if (modelUrlMap.has(groupValue) && !formulas[i][indices.SERIES_URL - 1]) {
          formulas[i][indices.SERIES_URL - 1] = createHyperlinkFormula(modelUrlMap.get(groupValue), groupValue);
          modified = true;
        }
      }
    });

    if (modified) {
      range.setFormulas(formulas);
    }
    ss.toast("資料フォルダの作成とリンク設定が完了しました。");
  } catch (error) {
    Logger.log(error.stack);
    ss.toast(`エラー: ${error.message}`);
  }
}

/**
 * テンプレートのスプレッドシートをコピーし、情報を書き込んでフォルダに保存する
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
    
    const newFile = templateFile.makeCopy(newFileName, targetFolder);
    const newSpreadsheet = SpreadsheetApp.openById(newFile.getId());
    const sheet = newSpreadsheet.getSheets()[0];

    // 指定のセルに情報を書き込む
    sheet.getRange("B4").setValue("機種：" + model); // B4セルに機種
    sheet.getRange("B5").setValue("製番：" + kiban); // B5セルに製番

    // シート全体の書式を設定
    sheet.getDataRange().setFontFamily("Arial").setFontSize(11);

    SpreadsheetApp.flush();

    Logger.log(`管理表「${newFileName}」をフォルダ「${targetFolder.getName()}」に作成しました。`);
  } catch (e) {
    Logger.log(`管理表の作成中にエラー: ${e.message}`);
    SpreadsheetApp.getActiveSpreadsheet().toast("管理表の作成に失敗しました。ログを確認してください。");
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