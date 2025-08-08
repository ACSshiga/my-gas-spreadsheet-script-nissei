/**
 * Code.gs
 * イベントハンドラ、カスタムメニュー、トリガー設定を管理する司令塔。
 * QUERY関数による常時同期ビューと、動的な条件付き書式に対応。
 */

// =================================================================================
// === イベントハンドラ ===
// =================================================================================

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー')
    .addItem('ソートビューを作成', 'createSortedView')
    .addItem('ソートビューを全て削除', 'removeAllSortedViews')
    .addSeparator()
    .addItem('請求シートを更新', 'showBillingSidebar')
    .addSeparator()
    .addItem('全資料フォルダ作成', 'bulkCreateMaterialFolders')
    .addItem('週次バックアップを作成', 'createWeeklyBackup')
    .addSeparator()
    .addItem('各種設定と書式を再適用', 'runAllManualMaintenance')
    .addItem('シート全体の書式を整える', 'applyStandardFormattingToAllSheets')
    .addItem('次の月のカレンダーを追加', 'addNextMonthColumnsToAllInputSheets')
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
      colorizeSheet_(new MainSheet());
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
// === ソートビュー（QUERY関数版）機能 ===
// =================================================================================

/**
 * メインシートのデータを常時同期・ソートする新しいシート「ソートビュー」を作成します。
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

  // 新しいビューシートの名前を決定（連番処理）
  let viewSheetName = 'ソートビュー';
  let counter = 2;
  while (ss.getSheetByName(viewSheetName)) {
    viewSheetName = `ソートビュー (${counter})`;
    counter++;
  }

  // QUERY関数を構築
  const mainSheetName = mainSheet.getName();
  const lastColLetter = mainSheet.getRange(1, mainSheet.getLastColumn()).getA1Notation().replace("1", "");
  const dataRange = `'${mainSheetName}'!A2:${lastColLetter}`; // ヘッダーを除いたデータ範囲
  const sortColLetter = String.fromCharCode(64 + sortColumnIndex);
  const formula = `=QUERY(${dataRange}, "SELECT * ORDER BY ${sortColLetter} ASC")`;

  // シートを作成
  const viewSheet = ss.insertSheet(viewSheetName, 0);
  
  // ヘッダーをコピー
  const headers = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn());
  headers.copyTo(viewSheet.getRange(1, 1));

  // A2セルにQUERY関数をセット
  viewSheet.getRange("A2").setFormula(formula);

  // 少し待ってから後続処理を実行（QUERYの結果が反映されるのを待つ）
  Utilities.sleep(1500);

  // フィルターを作成
  const dataRangeWithHeader = viewSheet.getRange(1, 1, viewSheet.getLastRow(), viewSheet.getLastColumn());
  dataRangeWithHeader.createFilter();

  // 書式と条件付き書式を適用
  applyStandardFormattingToMainSheet();
  applyConditionalFormattingToView(viewSheet);
  
  viewSheet.activate();
  ss.toast(`${viewSheetName} を作成しました。`, '完了', 5);
}

/**
 * すべてのソートビューを削除します。
 */
function removeAllSortedViews(showMessage = true) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(sheet => {
    if (sheet.getName().startsWith('ソートビュー')) {
      ss.deleteSheet(sheet);
    }
  });

  const mainSheet = ss.getSheetByName(CONFIG.SHEETS.MAIN)
  if(mainSheet) mainSheet.activate();
  if (showMessage) {
    ss.toast('すべてのソートビューを削除しました。', '完了', 3);
  }
}

/**
 * ビューシートに条件付き書式（色付け）を適用します。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象のビューシート
 */
function applyConditionalFormattingToView(sheet) {
  const indices = getColumnIndices(sheet, MAIN_SHEET_HEADERS);
  const lastRow = sheet.getMaxRows();
  const rules = [];

  // 交互の行の背景色
  const alternateRowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=MOD(ROW(), 2) = 0')
    .setBackground(CONFIG.COLORS.ALTERNATE_ROW)
    .setRanges([sheet.getRange(2, 1, lastRow -1, sheet.getLastColumn())])
    .build();
  rules.push(alternateRowRule);

  const colorMasters = [
    { master: CONFIG.SHEETS.SHINCHOKU_MASTER, colIndex: indices.PROGRESS, colorCol: 1 },
    { master: CONFIG.SHEETS.TANTOUSHA_MASTER, colIndex: indices.TANTOUSHA, colorCol: 2 },
    { master: CONFIG.SHEETS.SAGYOU_KUBUN_MASTER, colIndex: indices.SAGYOU_KUBUN, colorCol: 1 },
    { master: CONFIG.SHEETS.TOIAWASE_MASTER, colIndex: indices.TOIAWASE, colorCol: 1 }
  ];

  colorMasters.forEach(item => {
    if (!item.colIndex) return;
    const colorMap = getColorMapFromMaster(item.master, 0, item.colorCol);
    const range = sheet.getRange(2, item.colIndex, lastRow - 1, 1);
    
    for (const [key, color] of colorMap.entries()) {
      if (key && color) {
        const rule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(key)
          .setBackground(color)
          .setRanges([range])
          .build();
        rules.push(rule);
      }
    }
  });

  sheet.setConditionalFormatRules(rules);
}


// =================================================================================
// === 書式設定、色付け、その他（既存の関数群） ===
// =================================================================================

// onOpenとonEdit以外の既存の関数は変更がないため、ここでは省略します。
// 以下の関数群は、以前のコードのままご利用ください。
//
// periodicMaintenance()
// runAllManualMaintenance()
// applyStandardFormattingToAllSheets()
// applyStandardFormattingToMainSheet()
// setupAllDataValidations()
// colorizeAllSheets()
// colorizeSheet_()
// ... その他の既存ヘルパー関数

// ▼▼▼▼▼ 以下、既存の関数（変更なし） ▼▼▼▼▼

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
      if(sheet.getName().startsWith('ソートビュー')) return; // ソートビューは別途処理
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

function colorizeAllSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();

    allSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      try {
        if (sheetName === CONFIG.SHEETS.MAIN) {
          colorizeSheet_(new MainSheet());
        } else if (sheetName.startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
          const tantoushaName = sheetName.replace(CONFIG.SHEETS.INPUT_PREFIX, '');
          colorizeSheet_(new InputSheet(tantoushaName));
        }
      } catch (e) {
        Logger.log(`シート「${sheetName}」の色付け処理中にエラー: ${e.message}`);
      }
    });
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
  if (lastCol === 0) return;
  const fullRange = sheet.getRange(startRow, 1, dataRows, lastCol);

  const values = fullRange.getValues();
  const formulas = fullRange.getFormulasR1C1(); // Use R1C1 for robust formula handling
  const backgroundColors = fullRange.getBackgrounds();

  const mgmtNoCol = indices.MGMT_NO;
  const progressCol = indices.PROGRESS;
  const tantoushaCol = indices.TANTOUSHA;
  const toiawaseCol = indices.TOIAWASE;
  const kibanCol = indices.KIBAN;
  const sagyouKubunCol = indices.SAGYOU_KUBUN;
  const DUPLICATE_COLOR = '#cccccc';
  const uniqueKeys = new Set();
  
  values.forEach((row, i) => {
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
    
    // 背景色のリセット
    const baseColor = (i % 2 !== 0) ? CONFIG.COLORS.ALTERNATE_ROW : CONFIG.COLORS.DEFAULT_BACKGROUND;
    for (let j = 0; j < lastCol; j++) {
       backgroundColors[i][j] = baseColor;
    }
    
    if (isDuplicate) {
      for (let j = 0; j < lastCol; j++) {
        backgroundColors[i][j] = DUPLICATE_COLOR;
      }
    } else {
      if (progressCol) {
        const progressValue = safeTrim(row[progressCol - 1]);
        const progressColor = getColor(PROGRESS_COLORS, progressValue, baseColor);
        backgroundColors[i][progressCol - 1] = progressColor;
        if (mgmtNoCol) {
          backgroundColors[i][mgmtNoCol - 1] = progressColor;
        }
      }

      if (sagyouKubunCol) {
        backgroundColors[i][sagyouKubunCol - 1] = getColor(SAGYOU_KUBUN_COLORS, safeTrim(row[sagyouKubunCol - 1]), baseColor);
      }

      if (sheetObject instanceof MainSheet) {
        if (tantoushaCol) {
          backgroundColors[i][tantoushaCol - 1] = getColor(TANTOUSHA_COLORS, safeTrim(row[tantoushaCol - 1]), baseColor);
        }
        if (toiawaseCol) {
          backgroundColors[i][toiawaseCol - 1] = getColor(TOIAWASE_COLORS, safeTrim(row[toiawaseCol - 1]), baseColor);
        }
      }
    }
  });

  fullRange.setBackgrounds(backgroundColors);
}