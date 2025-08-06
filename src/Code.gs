/**
 * Code.gs
 * イベントハンドラ、カスタムメニュー、トリガー設定を管理する司令塔。
 */

// =================================================================================
// === イベントハンドラ ===
// =================================================================================

/**
 * onOpen: ファイルを開いた時に実行される処理
 */
function onOpen(e) {
  // カスタムメニューの作成
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー')
    .addItem('自分の担当案件のみ表示', 'createPersonalView')
    .addItem('表示を更新', 'refreshPersonalView') 
    .addItem('フィルタ表示を終了', 'removePersonalView')
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

  // onOpenで実行する処理の順番を最適化
  syncMainToAllInputSheets();
  applyStandardFormattingToAllSheets();
  applyStandardFormattingToMainSheet();
  colorizeAllSheets();
}

/**
 * onEdit: 編集イベントを処理する司令塔 (待機処理を追加)
 */
function onEdit(e) {
  if (!e || !e.source || !e.range) return;
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); 
  } catch (err) {
    Logger.log('ロックの待機中にタイムアウトしました。');
    SpreadsheetApp.getActiveSpreadsheet().toast('他のユーザーの処理が長引いているため、今回の編集は反映されませんでした。時間をおいて再度お試しください。');
    return;
  }
  
  const ss = e.source;
  try {
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();

    if (sheetName === CONFIG.SHEETS.MAIN) {
      const mainSheet = new MainSheet();
      const editedCol = e.range.getColumn();
      if (editedCol === mainSheet.indices.TANTOUSHA) {
        ss.toast('担当者の変更を検出しました。関連シートを同期します...', '同期中', 5);
        syncMainToAllInputSheets();
        colorizeAllSheets();
        ss.toast('同期処理が完了しました。', '完了', 3);
      } else {
         colorizeAllSheets();
      }
    } else if (sheetName.startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
      ss.toast(`${sheetName}の変更を検出しました。メインシートへ集計します...`, '集計中', 5);
      syncInputToMain(sheetName, e.range);
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

// =================================================================================
// === メンテナンス用関数 ===
// =================================================================================

/**
 * 定期実行（毎時）で呼び出される関数
 */
function periodicMaintenance() {
  setupAllDataValidations();
}

/**
 * 手動実行用の統合関数
 */
function runAllManualMaintenance() {
  SpreadsheetApp.getActiveSpreadsheet().toast('各種設定と書式を適用中...', '処理中', 3);
  applyStandardFormattingToAllSheets();
  applyStandardFormattingToMainSheet();
  colorizeAllSheets();
  setupAllDataValidations();
  SpreadsheetApp.getActiveSpreadsheet().toast('適用が完了しました。', '完了', 3);
}

// =================================================================================
// === 書式設定 ===
// =================================================================================

/**
 * 全てのシートにおしゃれな書式を適用します。
 */
function applyStandardFormattingToAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  
  ss.toast('全シートの書式をおしゃれに更新中...', '処理中');
  
  allSheets.forEach(sheet => {
    try {
      const dataRange = sheet.getDataRange();
      if (dataRange.isBlank()) return;
      const lastCol = sheet.getLastColumn();
      const lastRow = sheet.getLastRow();

      // 1. 基本書式（フォント、サイズ、配置、罫線）を適用
      dataRange
        .setFontFamily("Arial")
        .setFontSize(12)
        .setVerticalAlignment("middle")
        .setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);

      // 2. ヘッダーを中央揃え
      const headerRange = sheet.getRange(1, 1, 1, lastCol);
      headerRange.setHorizontalAlignment("center");

      // 3. 数字の列を右揃え
      const sheetName = sheet.getName();
      let indices;
      if (sheetName === CONFIG.SHEETS.MAIN || sheetName.startsWith('View_')) {
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
           const dateColStart = indices.SEPARATOR + 2; // 日付入力開始列
           if (lastCol >= dateColStart && lastRow > 2) {
             sheet.getRange(3, dateColStart, lastRow - 2, lastCol - dateColStart + 1).setHorizontalAlignment("right");
           }
        }
      }
    } catch (e) {
      Logger.log(`シート「${sheet.getName()}」の書式設定中にエラー: ${e.message}`);
    }
  });
  
  ss.toast('全シートの書式を更新しました。', '完了', 3);
}


/**
 * メインシートとViewシートにヘッダーの色付けとウィンドウ枠の固定を適用します。
 */
function applyStandardFormattingToMainSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToFormat = [ss.getSheetByName(CONFIG.SHEETS.MAIN), ss.getActiveSheet()];
  
  sheetsToFormat.forEach(sheet => {
    if (sheet && (sheet.getName() === CONFIG.SHEETS.MAIN || sheet.getName().startsWith('View_'))) {
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
  
  const mainDataRange = mainSheet.getDataRange();
  const mainValues = mainDataRange.getValues();
  const mainFormulas = mainDataRange.getFormulas();
  
  const headers = mainFormulas[0]; // ヘッダー行は数式ごと取得
  const personalData = [headers];

  mainValues.forEach((row, index) => {
    if (index > 0 && row[mainIndices.TANTOUSHA - 1] === tantoushaName) {
      const rowData = row.map((cell, colIndex) => mainFormulas[index][colIndex] || cell);
      personalData.push(rowData);
    }
  });

  removePersonalView(false);

  const viewSheetName = `View_${tantoushaName}`;
  const viewSheet = ss.insertSheet(viewSheetName, 0);
  
  viewSheet.getRange(1, 1, personalData.length, headers.length).setValues(personalData);
  
  applyStandardFormattingToAllSheets();
  applyStandardFormattingToMainSheet();
  
  if (personalData.length > 1) {
    viewSheet.getRange(1, 1, personalData.length, headers.length).createFilter();
  }
  
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

// =================================================================================
// === 色付け処理 ===
// =================================================================================

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
        } else if (sheetName.startsWith('View_')) {
          const viewSheetObj = {
            getSheet: () => sheet,
            indices: getColumnIndices(sheet, MAIN_SHEET_HEADERS),
            startRow: 2,
            getLastRow: () => sheet.getLastRow()
          };
          colorizeSheet_(viewSheetObj);
        }
      } catch (e) {
        Logger.log(`シート「${sheetName}」の色付け処理中にエラー: ${e.message}`);
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
  if (lastCol === 0) return;
  const fullRange = sheet.getRange(startRow, 1, dataRows, lastCol);

  const values = fullRange.getValues();
  const formulas = fullRange.getFormulas();
  const outputValues = values.map((row, i) => {
    return row.map((cell, j) => formulas[i][j] || cell);
  });

  const backgroundColors = fullRange.getBackgrounds();

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
      const baseColor = (i % 2 !== 0) ? CONFIG.COLORS.ALTERNATE_ROW : CONFIG.COLORS.DEFAULT_BACKGROUND;
      for (let j = 0; j < lastCol; j++) {
        if (!formulas[i][j]) {
          backgroundColors[i][j] = baseColor;
        }
      }

      if (progressCol) {
        const progressValue = safeTrim(row[progressCol - 1]);
        const progressColor = (progressValue === "") ? baseColor : getColor(PROGRESS_COLORS, progressValue, baseColor);
        
        backgroundColors[i][progressCol - 1] = progressColor;
        if (mgmtNoCol) {
          backgroundColors[i][mgmtNoCol - 1] = progressColor;
        }
      }

      if (sagyouKubunCol) {
        const color = getColor(SAGYOU_KUBUN_COLORS, safeTrim(row[sagyouKubunCol - 1]), baseColor);
        if (color !== baseColor) {
          backgroundColors[i][sagyouKubunCol - 1] = color;
        }
      }

      if (sheetObject instanceof MainSheet || sheet.getName().startsWith('View_')) {
        if (tantoushaCol) {
          const color = getColor(TANTOUSHA_COLORS, safeTrim(row[tantoushaCol - 1]), baseColor);
          if (color !== baseColor) {
            backgroundColors[i][tantoushaCol - 1] = color;
          }
        }
        if (toiawaseCol) {
          const color = getColor(TOIAWASE_COLORS, safeTrim(row[toiawaseCol - 1]), baseColor);
          if (color !== baseColor) {
            backgroundColors[i][toiawaseCol - 1] = color;
          }
        }
      }
      if (sheetObject instanceof MainSheet && tantoushaCol) {
        normalRanges.push(sheet.getRange(startRow + i, tantoushaCol));
      }
    }
  });

  fullRange.setValues(outputValues);
  fullRange.setBackgrounds(backgroundColors);

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