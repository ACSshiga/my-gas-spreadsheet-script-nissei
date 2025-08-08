/**
 * Code.gs
 * イベントハンドラ、カスタムメニュー、トリガー設定を管理する司令塔。
 * データコピー方式による、リンク・書式対応の多機能ソートビューを実装。(最終改訂版)
 */

// =================================================================================
// === イベントハンドラ ===
// =================================================================================

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー')
    .addItem('ソートビューを作成', 'createSortedView')
    .addItem('表示を更新', 'refreshSortedView') // 更新機能を明確にメニューに追加
    .addItem('ソートビューを全て削除', 'removeAllSortedViews')
    .addSeparator()
    .addItem('請求シートを更新', 'showBillingSidebar')
    .addSeparator()
    .addItem('全資料フォルダ作成', 'bulkCreateMaterialFolders')
    .addItem('週次バックアップを作成', 'createWeeklyBackup')
    .addSeparator()
    .addItem('各種設定と書式を再適用', 'runAllManualMaintenance')
    .addSeparator()
    .addItem('スクリプトのキャッシュをクリア', 'clearScriptCache')
    .addItem('フォルダからインポートを実行', 'importFromDriveFolder')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.source || !e.range) return;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); 
  } catch (err) {
    Logger.log('ロックの待機中にタイムアウトしました。');
    return;
  }
  
  const ss = e.source;
  try {
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();

    if (sheetName === CONFIG.SHEETS.MAIN) {
      colorizeSheet_(new MainSheet()); // メインシートの色付けのみ実行
      const mainSheet = new MainSheet();
      if (e.range.getColumn() === mainSheet.indices.TANTOUSHA) {
        syncMainToAllInputSheets();
      }
    } else if (sheetName.startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
      syncInputToMain(sheetName, e.range);
      colorizeSheet_(new InputSheet(sheetName.replace(CONFIG.SHEETS.INPUT_PREFIX, '')));
    }
  } catch (error) {
    Logger.log(error.stack);
    ss.toast(`エラーが発生しました: ${error.message}`, "エラー", 10);
  } finally {
    lock.releaseLock();
  }
}

// =================================================================================
// === ソートビュー（データコピー方式）機能 ===
// =================================================================================

/**
 * ソートビューを更新します。古いビューを全て削除し、新しいビューを作成します。
 */
function refreshSortedView() {
    SpreadsheetApp.getActiveSpreadsheet().toast('ビューを更新しています...', '処理中', 5);
    removeAllSortedViews(false); // メッセージなしで全ビューを削除
    createSortedView();
}

/**
 * メインシートのデータを数式ごとコピー・ソートし、新しい「ソートビュー」シートを作成します。
 */
function createSortedView() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(CONFIG.SHEETS.MAIN);
  if (!mainSheet) {
    SpreadsheetApp.getUi().alert('メインシートが見つかりません。');
    return;
  }

  const mainIndices = getColumnIndices(mainSheet, MAIN_SHEET_HEADERS);
  const sortColumnIndex = mainIndices.MGMT_NO;
  if (!sortColumnIndex) {
      SpreadsheetApp.getUi().alert('ソートキーとなる「管理No」列が見つかりません。');
      return;
  }

  const lastRow = mainSheet.getLastRow();
  const dataStartRow = CONFIG.DATA_START_ROW.MAIN;
  if (lastRow < dataStartRow) {
    SpreadsheetApp.getUi().alert('メインシートにデータがありません。');
    return;
  }

  // 新しいビューシートの名前を決定（連番処理）
  let viewSheetName = 'ソートビュー';
  let counter = 2;
  while (ss.getSheetByName(viewSheetName)) {
    viewSheetName = `ソートビュー (${counter})`;
    counter++;
  }
  
  const viewSheet = ss.insertSheet(viewSheetName, 0);

  // ★★★★★ 修正点：copyToで数式ごと完全コピー ★★★★★
  const sourceRange = mainSheet.getRange(1, 1, lastRow, mainSheet.getLastColumn());
  sourceRange.copyTo(viewSheet.getRange(1, 1), {contentsOnly: false}); // contentsOnly:falseで数式や書式もコピー
  // ★★★★★ 修正完了 ★★★★★

  // コピーしたデータ範囲に対して並べ替えを実行
  const viewRangeToSort = viewSheet.getRange(dataStartRow, 1, viewSheet.getLastRow() - (dataStartRow - 1), viewSheet.getLastColumn());
  viewRangeToSort.sort({column: sortColumnIndex, ascending: true});

  // フィルターと書式を適用
  viewSheet.getRange(1, 1, viewSheet.getLastRow(), viewSheet.getLastColumn()).createFilter();
  applyStandardFormattingToMainSheet(); // ヘッダー書式とウィンドウ枠固定
  
  // 色付け処理
  colorizeSheet_(new ViewSheet(viewSheet));

  viewSheet.activate();
  ss.toast(`${viewSheetName} を作成しました。`, '完了', 5);
}

/**
 * すべてのソートビューを削除します。
 */
function removeAllSortedViews(showMessage = true) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let deleted = false;
  ss.getSheets().forEach(sheet => {
    if (sheet.getName().startsWith('ソートビュー')) {
      ss.deleteSheet(sheet);
      deleted = true;
    }
  });

  const mainSheet = ss.getSheetByName(CONFIG.SHEETS.MAIN);
  if (mainSheet) mainSheet.activate();
  
  if (showMessage && deleted) {
    ss.toast('すべてのソートビューを削除しました。', '完了', 3);
  }
}

// =================================================================================
// === 色付け処理とヘルパークラス ===
// =================================================================================

/**
 * ソートビューシートを扱うための簡易クラス
 */
class ViewSheet {
  constructor(sheet) {
    this.sheet = sheet;
    this.startRow = CONFIG.DATA_START_ROW.MAIN;
    this.indices = getColumnIndices(this.sheet, MAIN_SHEET_HEADERS);
  }
  getSheet() { return this.sheet; }
  getLastRow() { return this.sheet.getLastRow(); }
}


function colorizeAllSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.getSheets().forEach(sheet => {
      const sheetName = sheet.getName();
      try {
        if (sheetName === CONFIG.SHEETS.MAIN) {
          colorizeSheet_(new MainSheet());
        } else if (sheetName.startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
          const tantoushaName = sheetName.replace(CONFIG.SHEETS.INPUT_PREFIX, '');
          colorizeSheet_(new InputSheet(tantoushaName));
        } else if (sheetName.startsWith('ソートビュー')) {
          colorizeSheet_(new ViewSheet(sheet));
        }
      } catch (e) {
        Logger.log(`シート「${sheetName}」の色付け処理中にエラー: ${e.message}`);
      }
    });
  } catch (error) {
    Logger.log(`色付け処理でエラーが発生しました: ${error.stack}`);
  }
}


/**
 * ★★★★★ 修正点：この関数は背景色のみを設定し、数式を絶対に上書きしない ★★★★★
 */
function colorizeSheet_(sheetObject) {
  const sheet = sheetObject.getSheet();
  const startRow = sheetObject.startRow;
  const lastRow = sheet.getLastRow();

  if (lastRow < startRow) return;

  const indices = sheetObject.indices;
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol);
  
  const values = range.getValues();
  const backgroundColors = []; // 新しい背景色配列を初期化

  const PROGRESS_COLORS = getColorMapFromMaster(CONFIG.SHEETS.SHINCHOKU_MASTER, 0, 1);
  const TANTOUSHA_COLORS = getColorMapFromMaster(CONFIG.SHEETS.TANTOUSHA_MASTER, 0, 2);
  const SAGYOU_KUBUN_COLORS = getColorMapFromMaster(CONFIG.SHEETS.SAGYOU_KUBUN_MASTER, 0, 1);
  const TOIAWASE_COLORS = getColorMapFromMaster(CONFIG.SHEETS.TOIAWASE_MASTER, 0, 1);

  const mgmtNoCol = indices.MGMT_NO;
  const progressCol = indices.PROGRESS;
  const tantoushaCol = indices.TANTOUSHA;
  const toiawaseCol = indices.TOIAWASE;
  const sagyouKubunCol = indices.SAGYOU_KUBUN;

  values.forEach((row, i) => {
    const rowColors = [];
    const baseColor = (i % 2 === 0) ? CONFIG.COLORS.DEFAULT_BACKGROUND : CONFIG.COLORS.ALTERNATE_ROW;
    
    // 全列の基本色を設定
    for (let j = 0; j < lastCol; j++) {
      rowColors[j] = baseColor;
    }

    // 進捗に応じた色付け (管理Noにも適用)
    if (progressCol) {
      const progressValue = safeTrim(row[progressCol - 1]);
      const progressColor = getColor(PROGRESS_COLORS, progressValue, baseColor);
      rowColors[progressCol - 1] = progressColor;
      if (mgmtNoCol) {
        rowColors[mgmtNoCol - 1] = progressColor;
      }
    }

    // 各マスターに応じた色付け
    if (sagyouKubunCol) {
      const value = safeTrim(row[sagyouKubunCol - 1]);
      const color = getColor(SAGYOU_KUBUN_COLORS, value, baseColor);
      if (color !== baseColor) rowColors[sagyouKubunCol - 1] = color;
    }
    if (tantoushaCol) {
       const value = safeTrim(row[tantoushaCol - 1]);
       const color = getColor(TANTOUSHA_COLORS, value, baseColor);
       if (color !== baseColor) rowColors[tantoushaCol - 1] = color;
    }
    if (toiawaseCol) {
       const value = safeTrim(row[toiawaseCol - 1]);
       const color = getColor(TOIAWASE_COLORS, value, baseColor);
       if (color !== baseColor) rowColors[toiawaseCol - 1] = color;
    }
    backgroundColors.push(rowColors);
  });

  // 背景色のみを更新する
  range.setBackgrounds(backgroundColors);
}

// =================================================================================
// === 既存のヘルパー関数群（変更なし） ===
// =================================================================================
// ... 以下のコードは変更がないため、そのまま貼り付けてください ...
function periodicMaintenance() {
  setupAllDataValidations();
}

function runAllManualMaintenance() {
  SpreadsheetApp.getActiveSpreadsheet().toast('各種設定と書式を適用中...', '処理中', 3);
  applyStandardFormattingToAllSheets();
  applyStandardFormattingToMainSheet();
  colorizeAllSheets();
  setupAllDataValidations();
  SpreadsheetApp.getActiveSpreadsheet().toast('適用が完了しました。', '完了', 3);
}

function applyStandardFormattingToAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  
  allSheets.forEach(sheet => {
    try {
      if(sheet.getName().startsWith('ソートビュー')) return; 
      const dataRange = sheet.getDataRange();
      if (dataRange.isBlank()) return;
      const lastCol = sheet.getLastColumn();
      const lastRow = sheet.getLastRow();

      dataRange
        .setFontFamily("Arial")
        .setFontSize(12)
        .setVerticalAlignment("middle")
        .setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);

      const headerRange = sheet.getRange(1, 1, 1, lastCol);
      headerRange.setHorizontalAlignment("center");

      const sheetName = sheet.getName();
      let indices;
      if (sheetName === CONFIG.SHEETS.MAIN) {
        indices = getColumnIndices(sheet, MAIN_SHEET_HEADERS);
        const numberCols = [indices.PLANNED_HOURS, indices.ACTUAL_HOURS];
        numberCols.forEach(colIndex => {
          if (colIndex && lastRow > 1) {
            sheet.getRange(2, colIndex, lastRow - 1).setHorizontalAlignment("right");
          }
        });
      } else if (sheetName.startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
        indices = getColumnIndices(sheet, INPUT_SHEET_HEADERS);
        const numberCols = [indices.PLANNED_HOURS, indices.ACTUAL_HOURS_SUM];
        numberCols.forEach(colIndex => {
          if (colIndex && lastRow > 2) {
             sheet.getRange(3, colIndex, lastRow - 2).setHorizontalAlignment("right");
          }
        });
        if (indices.SEPARATOR) {
           const dateColStart = indices.SEPARATOR + 2;
           if (lastCol >= dateColStart && lastRow > 2) {
             sheet.getRange(3, dateColStart, lastRow - 2, lastCol - dateColStart + 1).setHorizontalAlignment("right");
           }
        }
      }
    } catch (e) {
      Logger.log(`シート「${sheet.getName()}」の書式設定中にエラー: ${e.message}`);
    }
  });
}

function applyStandardFormattingToMainSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();

  allSheets.forEach(sheet => {
    if (sheet && (sheet.getName() === CONFIG.SHEETS.MAIN || sheet.getName().startsWith('ソートビュー'))) {
      try {
        const headerRange = sheet.getRange(1, 1, 1, sheet.getMaxColumns());
        headerRange.setBackground(CONFIG.COLORS.HEADER_BACKGROUND)
                   .setFontColor('#ffffff')
                   .setFontWeight('bold');
        
        sheet.setFrozenRows(1);
        sheet.setFrozenColumns(4);
      } catch(e) {
        Logger.log(`シート「${sheet.getName()}」のヘッダー書式設定中にエラー: ${e.message}`);
      }
    }
  });
}

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

            if(progressCol && lastRow >= inputSheet.startRow) {
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