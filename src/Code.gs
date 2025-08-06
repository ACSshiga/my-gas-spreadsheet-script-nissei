/**
 * Code.gs
 * イベントハンドラとカスタムメニューを管理する司令塔。
 */

// =================================================================================
// === イベントハンドラ ===
// =================================================================================

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('カスタムメニュー');
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  if (activeSheet.getName() === CONFIG.SHEETS.MAIN || activeSheet.getName().startsWith('View_')) {
    menu.addItem('自分の担当案件のみ表示', 'createPersonalView');
    menu.addItem('表示を更新', 'refreshPersonalView') 
    menu.addItem('フィルタ表示を終了', 'removePersonalView');
  } 
  else if (activeSheet.getName().startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
    menu.addItem('すべての月を表示', 'showAllDateColumns');
    menu.addItem('表示を当月・前月に戻す', 'hideOldDateColumns');
  }

  menu.addSeparator()
    .addItem('請求シートを更新', 'showBillingSidebar')
    .addSeparator()
    .addItem('全資料フォルダ作成', 'bulkCreateMaterialFolders')
    .addItem('週次バックアップを作成', 'createWeeklyBackup')
    .addSeparator()
    .addItem('重複チェックと色付けを再実行', 'runColorizeAllSheets')
    .addItem('スクリプトのキャッシュをクリア', 'clearScriptCache')
    .addSeparator()
    .addItem('フォルダからインポートを実行', 'importFromDriveFolder')
    .addToUi();
  
  setupAllDataValidations();
  syncDefaultProgressToMain();
  colorizeAllSheets();
}

/**
 * 申請書ファイルアップロード用のダイアログを表示します。
 */
function showImportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ImportFileDialog')
      .setWidth(400)
      .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, '申請書ファイルのインポート');
}


function onEdit(e) {
  if (!e || !e.source || !e.range) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    console.log('先行する処理が実行中のため、今回の編集イベントはスキップされました。');
    return;
  }
  
  const ss = e.source;

  try {
    setupAllDataValidations(); 

    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    if (sheetName === CONFIG.SHEETS.MAIN) {
      ss.toast('メインシートの変更を検出しました。同期処理を開始します...', '同期中', 5);
      syncMainToAllInputSheets();
      syncDefaultProgressToMain();
      colorizeAllSheets();
      ss.toast('同期処理が完了しました。', '完了', 3);
    } else if (sheetName.startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
      ss.toast(`${sheetName}の変更を検出しました。集計処理を開始します...`, '集計中', 5);
      syncInputToMain(sheetName, e.range);
      syncDefaultProgressToMain();
      colorizeAllSheets();
      ss.toast('集計処理が完了しました。', '完了', 3);
    }
  } catch (error) {
    Logger.log(error.stack);
    ss.toast(`エラーが発生しました: ${error.message}`, "エラー", 10);
  } finally {
    lock.releaseLock();
  }
}

function runColorizeAllSheets() {
  SpreadsheetApp.getActiveSpreadsheet().toast('重複チェックと色付けを実行中...', '処理中', 5);
  colorizeAllSheets();
  SpreadsheetApp.getActiveSpreadsheet().toast('処理が完了しました。', '完了', 3);
}


// =================================================================================
// === 工数シート表示切替機能 ===
// =================================================================================
function showAllDateColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet.getName().startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
    SpreadsheetApp.getUi().alert('この機能は工数シートでのみ利用できます。');
    return;
  }
  
  const lastCol = sheet.getLastColumn();
  const dateStartCol = Object.keys(INPUT_SHEET_HEADERS).length + 1;
  if (lastCol >= dateStartCol) {
    sheet.showColumns(dateStartCol, lastCol - dateStartCol + 1);
    SpreadsheetApp.getActiveSpreadsheet().toast('すべての日付列を表示しました。');
  }
}

function hideOldDateColumns() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (!sheet.getName().startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
      SpreadsheetApp.getUi().alert('この機能は工数シートでのみ利用できます。');
      return;
    }
    const tantoushaName = sheet.getName().replace(CONFIG.SHEETS.INPUT_PREFIX, '');
    const inputSheet = new InputSheet(tantoushaName);
    inputSheet.filterDateColumns();
    SpreadsheetApp.getActiveSpreadsheet().toast('表示を当月・前月に戻しました。');
}


// =================================================================================
// === 個人用ビュー（仮想シート）機能 ===
// =================================================================================
function refreshPersonalView() {
  SpreadsheetApp.getActiveSpreadsheet().toast('表示を更新しています...', '処理中', 5);
  removePersonalView(false);
  createPersonalView();
}

function createPersonalView() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userEmail = Session.getActiveUser().getEmail();
  const tantoushaName = getTantoushaNameByEmail(userEmail);
  if (!tantoushaName) {
    SpreadsheetApp.getUi().alert('あなたのメールアドレスが担当者マスタに見つかりません。');
    return;
  }

  const mainSheet = ss.getSheetByName(CONFIG.SHEETS.MAIN);
  const mainIndices = getColumnIndices(mainSheet, MAIN_SHEET_HEADERS);
  const mainData = mainSheet.getDataRange().getValues();
  
  const headers = mainData[0];
  const personalData = mainData.filter((row, index) => {
    return index === 0 || row[mainIndices.TANTOUSHA - 1] === tantoushaName;
  });
  removePersonalView(false);

  const viewSheetName = `View_${tantoushaName}`;
  const viewSheet = ss.insertSheet(viewSheetName, 0);
  viewSheet.getRange(1, 1, personalData.length, headers.length).setValues(personalData);
  if (personalData.length > 1) {
    viewSheet.getRange(1, 1, personalData.length, headers.length).createFilter();
  }
  
  viewSheet.autoResizeColumns(1, headers.length);
  try {
    const viewSheetObj = {
      getSheet: () => viewSheet,
      indices: mainIndices,
      startRow: 2,
      getLastRow: () => viewSheet.getLastRow()
    };
    colorizeSheet_(viewSheetObj);
  } catch (e) {
    Logger.log(`Viewシートの色付けエラー: ${e.message}`);
  }
  
  viewSheet.activate();
  ss.toast(`${tantoushaName}さん専用の表示を作成しました。`, '完了', 3);
}

function removePersonalView(showMessage = true) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userEmail = Session.getActiveUser().getEmail();
  const tantoushaName = getTantoushaNameByEmail(userEmail);
  if (tantoushaName) {
    const viewSheetName = `View_${tantoushaName}`;
    const sheetToDelete = ss.getSheetByName(viewSheetName);
    if (sheetToDelete) {
      ss.deleteSheet(sheetToDelete);
    }
  }

  const mainSheet = ss.getSheetByName(CONFIG.SHEETS.MAIN)
  if(mainSheet) mainSheet.activate();
  if (showMessage) {
    ss.toast('フィルタ表示を終了しました。', '完了', 3);
  }
}


// =================================================================================
// === データ入力規則の自動設定 ===
// =================================================================================
function setupAllDataValidations() {
  try {
    const mainSheetInstance = new MainSheet();
    const mainSheetObj = mainSheetInstance.getSheet(); 
    
    if (mainSheetObj.getLastRow() >= mainSheetInstance.startRow) {
      const mainLastRow = mainSheetObj.getMaxRows();
      const mainHeaderIndices = mainSheetInstance.indices;
      
      const mainValidationMap = {
        [CONFIG.SHEETS.SAGYOU_KUBUN_MASTER]: mainHeaderIndices.SAGYOU_KUBUN,
        [CONFIG.SHEETS.TOIAWASE_MASTER]: mainHeaderIndices.TOIAWASE,
        [CONFIG.SHEETS.TANTOUSHA_MASTER]: mainHeaderIndices.TANTOUSHA,
      };
      for (const [masterName, colIndex] of Object.entries(mainValidationMap)) {
        if(colIndex) {
          const masterValues = getMasterData(masterName).map(row => row[0]);
          if (masterValues.length > 0) {
            const rule = SpreadsheetApp.newDataValidation().requireValueInList(masterValues).setAllowInvalid(false).build();
            mainSheetObj.getRange(mainSheetInstance.startRow, colIndex, mainLastRow - mainSheetInstance.startRow + 1).setDataValidation(rule);
          }
        }
      }
      if (mainHeaderIndices.PROGRESS) {
        mainSheetObj.getRange(mainSheetInstance.startRow, mainHeaderIndices.PROGRESS, mainLastRow - mainSheetInstance.startRow + 1).clearDataValidations();
      }
    }

    const progressValues = getMasterData(CONFIG.SHEETS.SHINCHOKU_MASTER).map(row => row[0]);
    if (progressValues.length > 0) {
      const progressRule = SpreadsheetApp.newDataValidation().requireValueInList(progressValues).setAllowInvalid(false).build();
      const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
      allSheets.forEach(sheet => {
        if (sheet.getName().startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
          try {
            const inputSheet = new InputSheet(sheet.getName().replace(CONFIG.SHEETS.INPUT_PREFIX, ''));
            const lastRow = sheet.getMaxRows();
            const progressCol = inputSheet.indices.PROGRESS;

            if(progressCol && lastRow >= inputSheet.startRow) 
            {
              sheet.getRange(inputSheet.startRow, progressCol, lastRow - inputSheet.startRow + 1).setDataValidation(progressRule);
            }
          } catch(e) {
            Logger.log(`シート「${sheet.getName()}」の入力規則設定をスキップしました: ${e.message}`);
          }
        }
      });
    }
  } catch(e) {
    Logger.log(`データ入力規則の設定中にエラー: ${e.message}`);
  }
}

// =================================================================================
// === 色付け処理 ===
// =================================================================================

function colorizeAllSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const mainSheet = ss.getSheetByName(CONFIG.SHEETS.MAIN);
    if (mainSheet) {
      colorizeSheet_(new MainSheet());
    }

    const allSheets = ss.getSheets();
    allSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName.startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
        try {
          const tantoushaName = sheetName.replace(CONFIG.SHEETS.INPUT_PREFIX, '');
          colorizeSheet_(new InputSheet(tantoushaName));
        } catch (e) { /* エラーは無視 */ }
      } else if (sheetName.startsWith('View_')) {
        try {
          const viewSheetObj = {
            getSheet: () => sheet,
            indices: getColumnIndices(sheet, MAIN_SHEET_HEADERS),
            startRow: 2,
            getLastRow: () => sheet.getLastRow()
          };
          colorizeSheet_(viewSheetObj);
        } catch (e) {
          Logger.log(`Viewシート ${sheetName} の色付けエラー: ${e.message}`);
        }
      }
    });
    logWithTimestamp("全シートの色付け処理が完了しました。");
  } catch (error) {
    Logger.log(`色付け処理でエラーが発生しました: ${error.stack}`);
  }
}

function colorizeSheet_(sheetObject) {
  const PROGRESS_COLORS = getColorMapFromMaster(CONFIG.SHEETS.SHINCHOKU_MASTER, 0, 1);
  const TANTOUSHA_COLORS = getColorMapFromMaster(CONFIG.SHEETS.TANTOUSHA_MASTER, 0, 2);
  const SAGYOU_KUBUN_COLORS = getColorMapFromMaster(CONFIG.SHEETS.SAGYOU_KUBUN_MASTER, 0, 1);
  const TOIAWASE_COLORS = getColorMapFromMaster(CONFIG.SHEETS.TOIAWASE_MASTER, 0, 1);

  const sheet = sheetObject.getSheet();
  const indices = sheetObject.indices;
  const lastRow = sheetObject.getLastRow();
  const startRow = sheetObject.startRow;

  if (lastRow < startRow) return;
  const dataRows = lastRow - startRow + 1;
  const lastCol = sheet.getLastColumn();
  const fullRange = sheet.getRange(startRow, 1, dataRows, lastCol);
  const displayValues = fullRange.getDisplayValues();
  const formulas = fullRange.getFormulas();
  const backgroundColors = fullRange.getBackgrounds();
  
  const outputValues = JSON.parse(JSON.stringify(displayValues));

  const mgmtNoCol = indices.MGMT_NO;
  const progressCol = indices.PROGRESS;
  const tantoushaCol = indices.TANTOUSHA;
  const toiawaseCol = indices.TOIAWASE;
  const kibanCol = indices.KIBAN;
  const sagyouKubunCol = indices.SAGYOU_KUBUN;
  const DUPLICATE_COLOR = '#cccccc';

  const uniqueKeys = new Set();
  const restrictedRanges = [];
  const normalRanges = [];
  displayValues.forEach((row, i) => {
    let isDuplicate = false;
    if (kibanCol && sagyouKubunCol) {
      const kiban = safeTrim(row[kibanCol - 1]);
      const sagyouKubun = safeTrim(row[sagyouKubunCol - 1]);
      
      if (kiban && sagyouKubun) {
        const uniqueKey = `${kiban}_${sagyouKubun}`;
        if (uniqueKeys.has(uniqueKey)) {
          isDuplicate = true;
        } else {
          uniqueKeys.add(uniqueKey);
        }
      }
    }
    
    if (isDuplicate) {
      for (let j = 0; j < lastCol; j++) {
        backgroundColors[i][j] = DUPLICATE_COLOR;
      }
      if (progressCol) outputValues[i][progressCol - 1] = "機番重複";
      if (tantoushaCol) outputValues[i][tantoushaCol - 1] = "";
      if (sheetObject instanceof MainSheet && tantoushaCol) {
        restrictedRanges.push(sheet.getRange(startRow + i, tantoushaCol));
      }
    } else {
      for (let j = 0; j < lastCol; j++) {
          if(!formulas[i][j]) backgroundColors[i][j] = CONFIG.COLORS.DEFAULT_BACKGROUND;
      }

      if (progressCol) {
        const progressValue = safeTrim(row[progressCol - 1]);
        const progressColor = (progressValue === "")
          ? CONFIG.COLORS.DEFAULT_BACKGROUND
          : getColor(PROGRESS_COLORS, progressValue);
        
        backgroundColors[i][progressCol - 1] = progressColor;
        if (mgmtNoCol) backgroundColors[i][mgmtNoCol - 1] = progressColor;
      }

      if (sagyouKubunCol) {
        backgroundColors[i][sagyouKubunCol - 1] = getColor(SAGYOU_KUBUN_COLORS, safeTrim(row[sagyouKubunCol - 1]));
      }

      if (sheetObject instanceof MainSheet || sheet.getName().startsWith('View_')) {
        if (tantoushaCol) backgroundColors[i][tantoushaCol - 1] = getColor(TANTOUSHA_COLORS, safeTrim(row[tantoushaCol - 1]));
        if (toiawaseCol) backgroundColors[i][toiawaseCol - 1] = getColor(TOIAWASE_COLORS, safeTrim(row[toiawaseCol - 1]));
      }
      if (sheetObject instanceof MainSheet && tantoushaCol) {
        normalRanges.push(sheet.getRange(startRow + i, tantoushaCol));
      }
    }
  });

  formulas.forEach((row, i) => {
    row.forEach((formula, j) => {
      if (formula) {
        outputValues[i][j] = formula;
      }
    });
  });
  fullRange.setBackgrounds(backgroundColors);
  fullRange.setValues(outputValues);

  if (sheetObject instanceof MainSheet) {
    if (restrictedRanges.length > 0) {
      const restrictedRule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(sheet.getRange('A1:A1'), false)
        .setAllowInvalid(false)
        .setHelpText('この行は機番が重複しているため、担当者は設定できません。')
        .build();
      restrictedRanges.forEach(range => range.setDataValidation(restrictedRule));
    }
    if (normalRanges.length > 0) {
      const masterValues = getMasterData(CONFIG.SHEETS.TANTOUSHA_MASTER).map(row => row[0]);
      if (masterValues.length > 0) {
        const normalRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(masterValues)
          .setAllowInvalid(false)
          .build();
        normalRanges.forEach(range => range.setDataValidation(normalRule));
      } else {
         normalRanges.forEach(range => range.clearDataValidations());
      }
    }
  }
}