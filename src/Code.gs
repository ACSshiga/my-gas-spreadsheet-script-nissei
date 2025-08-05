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
    .addItem('表示を更新', 'refreshPersonalView') // ★更新ボタンを追加
    .addItem('フィルタ表示を終了', 'removePersonalView')
    .addSeparator()
    .addItem('請求シートを更新', 'showBillingSidebar')
    .addSeparator()
    .addItem('全機番の資料フォルダ作成', 'bulkCreateKibanFolders')
    .addItem('週次バックアップを作成', 'createWeeklyBackup')
    .addToUi();
  
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
function setupAllDataValidations() {
  try {
    const mainSheet = new MainSheet().getSheet();
    if (mainSheet.getLastRow() < CONFIG.DATA_START_ROW.MAIN) return;
    
    const lastRow = mainSheet.getMaxRows();
    const headerIndices = getColumnIndices(mainSheet, MAIN_SHEET_HEADERS);
    
    const validationMap = {
      [CONFIG.SHEETS.SAGYOU_KUBUN_MASTER]: headerIndices.SAGYOU_KUBUN,
      [CONFIG.SHEETS.SHINCHOKU_MASTER]: headerIndices.PROGRESS,
      [CONFIG.SHEETS.TOIAWASE_MASTER]: headerIndices.TOIAWASE,
      [CONFIG.SHEETS.TANTOUSHA_MASTER]: headerIndices.TANTOUSHA,
    };

    for (const [masterName, colIndex] of Object.entries(validationMap)) {
      if(colIndex) {
        const masterValues = getMasterData(masterName).flat();
        if (masterValues.length > 0) {
          const rule = SpreadsheetApp.newDataValidation().requireValueInList(masterValues).setAllowInvalid(false).build();
          mainSheet.getRange(CONFIG.DATA_START_ROW.MAIN, colIndex, lastRow - CONFIG.DATA_START_ROW.MAIN + 1).setDataValidation(rule);
        }
      }
    }
  } catch(e) {
    Logger.log(`データ入力規則の設定中にエラー: ${e.message}`);
  }
}