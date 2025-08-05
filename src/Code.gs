/**
 * Code.gs
 * イベントハンドラとカスタムメニューを管理する司令塔。
 * ★個人用の仮想シートを作成する方式に全面修正
 */

// =================================================================================
// === イベントハンドラ ===
// =================================================================================

function onOpen(e) {
  SpreadsheetApp.getUi().createMenu('カスタムメニュー')
    .addItem('自分の担当案件のみ表示', 'createPersonalView')
    .addItem('表示を更新', 'refreshPersonalView') 
    .addItem('フィルタ表示を終了', 'removePersonalView')
    .addSeparator()
    .addItem('請求シートを更新', 'showBillingSidebar')
    .addSeparator()
    .addItem('全機番の資料フォルダ作成', 'bulkCreateKibanFolders')
    .addItem('週次バックアップを作成', 'createWeeklyBackup')
    .addToUi();
  
  // ★起動時にメインシートと全工数シートの入力規則を更新
  setupAllDataValidations();
}

function onEdit(e) {
  // (DataSync.gsに処理が実装されているため、ここのロジックは変更なし)
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
// === ★個人用ビュー（仮想シート）機能 ===
// =================================================================================

/**
 * ★表示を更新するための関数（削除と作成を連続実行）
 */
function refreshPersonalView() {
  SpreadsheetApp.getActiveSpreadsheet().toast('表示を更新しています...', '処理中', 5);
  removePersonalView(false); // メッセージなしで削除
  createPersonalView();
}


/**
 * 現在のユーザー専用のビューシートを作成します。
 */
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

  // 既存の個人用シートがあれば削除
  removePersonalView(false); // メッセージなしで削除

  const viewSheetName = `View_${tantoushaName}`;
  const viewSheet = ss.insertSheet(viewSheetName, 0);
  viewSheet.getRange(1, 1, personalData.length, headers.length).setValues(personalData);
  
  viewSheet.getDataRange().createFilter(); 
  
  viewSheet.autoResizeColumns(1, headers.length);
  viewSheet.activate();
  
  ss.toast(`${tantoushaName}さん専用の表示を作成しました。`, '完了', 3);
}

/**
 * ユーザー専用のビューシートを削除します。
 * @param {boolean} [showMessage=true] - 完了メッセージを表示するかどうか
 */
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
  // メインシートをアクティブに戻す
  const mainSheet = ss.getSheetByName(CONFIG.SHEETS.MAIN)
  if(mainSheet) mainSheet.activate();
  
  if (showMessage) {
    ss.toast('フィルタ表示を終了しました。', '完了', 3);
  }
}


// =================================================================================
// === データ入力規則の自動設定 ===
// =================================================================================
/**
 * ★★★ 修正箇所 ★★★
 * メインシートと、すべての工数シートにドロップダウンリストを設定します。
 */
function setupAllDataValidations() {
  try {
    // 1. メインシートへの設定
    const mainSheet = new MainSheet();
    const mainSheetObj = mainSheet.getSheet();
    if (mainSheetObj.getLastRow() >= mainSheet.startRow) {
      const mainLastRow = mainSheetObj.getMaxRows();
      const mainHeaderIndices = getColumnIndices(mainSheetObj, MAIN_SHEET_HEADERS);
      
      const mainValidationMap = {
        [CONFIG.SHEETS.SAGYOU_KUBUN_MASTER]: mainHeaderIndices.SAGYOU_KUBUN,
        // [CONFIG.SHEETS.SHINCHOKU_MASTER]: mainHeaderIndices.PROGRESS, // ★メインシートの進捗ドロップダウンを削除
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
      // ★メインシートの進捗列の入力規則をクリア
      if (mainHeaderIndices.PROGRESS) {
        mainSheetObj.getRange(mainSheet.startRow, mainHeaderIndices.PROGRESS, mainLastRow - mainSheet.startRow + 1).clearDataValidations();
      }
    }

    // 2. 全工数シートへの設定
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