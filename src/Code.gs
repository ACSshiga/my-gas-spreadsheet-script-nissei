/**
 * Code.gs
 * イベントハンドラとカスタムメニューを管理する司令塔。
 * ★個人用の仮想シートを作成する方式に全面修正
 */

// =================================================================================
// === イベントハンドラ ===
// =================================================================================

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  // ★メニュー構成を更新
  const menu = ui.createMenu('カスタムメニュー');
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // メインシート用のメニュー
  if (activeSheet.getName() === CONFIG.SHEETS.MAIN) {
    menu.addItem('自分の担当案件のみ表示', 'createPersonalView');
    menu.addItem('フィルタ表示を終了', 'removePersonalView');
  } 
  // 工数シート用のメニュー
  else if (activeSheet.getName().startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
    menu.addItem('すべての月を表示', 'showAllDateColumns');
    menu.addItem('表示を当月・前月に戻す', 'hideOldDateColumns');
  }

  menu.addSeparator()
    .addItem('請求シートを更新', 'showBillingSidebar')
    .addSeparator()
    .addItem('全機番の資料フォルダ作成', 'bulkCreateKibanFolders')
    .addItem('週次バックアップを作成', 'createWeeklyBackup')
    .addToUi();
  
  // 起動時の初期化処理
  setupAllDataValidations();
}

function onEdit(e) {
  if (!e || !e.source || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const ss = e.source;

  try {
    if (sheetName === CONFIG.SHEETS.MAIN) {
      ss.toast('メインシートの変更を検出しました。同期処理を開始します...', '同期中', 5);
      syncMainToAllInputSheets();
      ss.toast('同期処理が完了しました。', '完了', 3);
    } else if (sheetName.startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
      ss.toast(`${sheetName}の変更を検出しました。集計処理を開始します...`, '集計中', 5);
      syncInputToMain(sheetName, e.range);
      ss.toast('集計処理が完了しました。', '完了', 3);
    }
  } catch (error) {
    Logger.log(error.stack);
    ss.toast(`エラーが発生しました: ${error.message}`, "エラー", 10);
  }
}

// =================================================================================
// === ★工数シート表示切替機能 ===
// =================================================================================
/**
 * 工数シートの非表示になっている日付列をすべて表示します。
 */
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

/**
 * 工数シートの表示を当月・前月に戻します（リロードと同じ効果）。
 */
function hideOldDateColumns() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (!sheet.getName().startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
      SpreadsheetApp.getUi().alert('この機能は工数シートでのみ利用できます。');
      return;
    }
    const tantoushaName = sheet.getName().replace(CONFIG.SHEETS.INPUT_PREFIX, '');
    const inputSheet = new InputSheet(tantoushaName);
    inputSheet.filterDateColumns(); // この関数が当月・前月以外を非表示にする
    SpreadsheetApp.getActiveSpreadsheet().toast('表示を当月・前月に戻しました。');
}


// =================================================================================
// === 個人用ビュー（仮想シート）機能 ===
// =================================================================================

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
  
  viewSheet.getDataRange().createFilter(); 
  
  viewSheet.autoResizeColumns(1, headers.length);
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
    const mainSheet = new MainSheet();
    const mainSheetObj = mainSheet.getSheet();
    if (mainSheetObj.getLastRow() >= mainSheet.startRow) {
      const mainLastRow = mainSheetObj.getMaxRows();
      const mainHeaderIndices = getColumnIndices(mainSheetObj, MAIN_SHEET_HEADERS);
      
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
            mainSheetObj.getRange(mainSheet.startRow, colIndex, mainLastRow - mainSheet.startRow + 1).setDataValidation(rule);
          }
        }
      }

      if (mainHeaderIndices.PROGRESS) {
        mainSheetObj.getRange(mainSheet.startRow, mainHeaderIndices.PROGRESS, mainLastRow - mainSheet.startRow + 1).clearDataValidations();
      }
    }

    const progressValues = getMasterData(CONFIG.SHEETS.SHINCHOKU_MASTER).flat();
    if (progressValues.length === 0) return;

    const progressRule = SpreadsheetApp.newDataValidation().requireValueInList(progressValues).setAllowInvalid(false).build();
    const tantoushaList = mainSheet.getTantoushaList();

    tantoushaList.forEach(tantousha => {
      try {
        const inputSheet = new InputSheet(tantousha.name);
        const inputSheetObj = inputSheet.getSheet();
        const inputHeaderIndices = getColumnIndices(inputSheetObj, INPUT_SHEET_HEADERS);
        const progressCol = inputHeaderIndices.PROGRESS;
        
        if (progressCol && inputSheetObj.getLastRow() >= inputSheet.startRow) {
          const inputLastRow = inputSheetObj.getMaxRows();
          inputSheetObj.getRange(inputSheet.startRow, progressCol, inputLastRow - inputSheet.startRow + 1).setDataValidation(progressRule);
        }
      } catch (e) { /* シートがなくてもエラーにしない */ }
    });

  } catch(e) {
    Logger.log(`データ入力規則の設定中にエラー: ${e.message}`);
  }
}