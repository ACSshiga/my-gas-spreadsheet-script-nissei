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
    .addToUi();
  setupAllDataValidations();
  
  syncDefaultProgressToMain();
  colorizeAllSheets();
}


/**
 * ★★★ 修正箇所 ★★★
 * 変数ssの宣言をtryブロックの外に出し、catchブロックでも使えるように修正。
 * これにより、エラー発生時にss is not definedというエラーが出るのを防ぎます。
 */
function onEdit(e) {
  if (!e || !e.source || !e.range) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    console.log('先行する処理が実行中のため、今回の編集イベントはスキップされました。');
    return;
  }
  
  const ss = e.source; // ★★★ ssの宣言をここに移動 ★★★

  try {
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

/**
 * デバッグ用：メニューから色付け処理だけを実行するための関数です。
 */
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
          const masterValues = getMasterData(masterName).flat();
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

    const progressValues = getMasterData(CONFIG.SHEETS.SHINCHOKU_MASTER).flat();
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