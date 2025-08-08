/**
 * Code.gs
 * イベントハンドラ、カスタムメニュー、トリガー設定を管理する司令塔。
 * QUERY関数による常時同期ビューに対応。
 */

// =================================================================================
// === イベントハンドラ ===
// =================================================================================

/**
 * onOpen: ファイルを開いた時に実行される処理
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー')
    .addItem('ソートビューを作成', 'createSortedView')
    .addItem('ソートビューを全て削除', 'removeAllSortedViews') // 変更
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

  // onOpenで実行する処理
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

// =================================================================================
// === 書式設定 ===
// =================================================================================

function applyStandardFormattingToAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  ss.toast('全シートの書式を更新中...', '処理中');
  
  allSheets.forEach(sheet => {
    try {
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
      if (sheetName === CONFIG.SHEETS.MAIN || sheetName.startsWith('ソートビュー')) {
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
  
  ss.toast('全シートの書式を更新しました。', '完了', 3);
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

// =================================================================================
// === ソートビュー（QUERY関数版）機能 ===
// =================================================================================

/**
 * メインシートのデータを常時同期・ソートする新しいシート「ソートビュー」を作成します。
 * 既に存在する場合は連番を付与して複数作成します。
 */
function createSortedView() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(CONFIG.SHEETS.MAIN);
  if (!mainSheet) {
    SpreadsheetApp.getUi().alert('メインシートが見つかりません。');
    return;
  }

  const mainIndices = getColumnIndices(mainSheet, MAIN_SHEET_HEADERS);
  const sortColumnIndex = mainIndices.MGMT_NO; // '管理No'でソート
  if (!sortColumnIndex) {
      SpreadsheetApp.getUi().alert('ソートキーとなる「管理No」列が見つかりません。');
      return;
  }

  // --- 新しいビューシートの名前を決定（連番処理） ---
  let viewSheetName = 'ソートビュー';
  let counter = 2;
  while (ss.getSheetByName(viewSheetName)) {
    viewSheetName = `ソートビュー (${counter})`;
    counter++;
  }

  // --- QUERY関数を構築 ---
  const mainSheetName = mainSheet.getName();
  const lastColLetter = mainSheet.getRange(1, mainSheet.getLastColumn()).getA1Notation().replace("1", "");
  const dataRange = `'${mainSheetName}'!A:${lastColLetter}`;
  
  // アルファベットの列名を取得 (1 -> A, 2 -> B, ...)
  const sortColLetter = String.fromCharCode(64 + sortColumnIndex);

  // メインシートのヘッダー行を除いてクエリを実行し、結果をソートする
  const formula = `=QUERY(${dataRange}, "SELECT * WHERE A IS NOT NULL ORDER BY ${sortColLetter} ASC", 1)`;

  // --- シートを作成し、関数をセット ---
  const viewSheet = ss.insertSheet(viewSheetName, 0);
  viewSheet.getRange("A1").setFormula(formula);

  // --- 書式を適用 ---
  applyStandardFormattingToMainSheet(); // ヘッダーと固定枠
  
  viewSheet.activate();
  ss.toast(`${viewSheetName} を作成しました。メインシートと常時同期されます。`, '完了', 5);
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
        } 
        // QUERY関数で生成されるビューの色付けは、元のメインシートの書式を引き継ぐため、
        // ここでの個別の色付け処理は不要になります。
        // もしビュー独自の書式ルールが必要な場合は、ここのコメントを解除して実装します。
        /*
        else if (sheetName.startsWith('ソートビュー')) {
          const viewSheetObj = {
            getSheet: () => sheet,
            indices: getColumnIndices(sheet, MAIN_SHEET_HEADERS),
            startRow: 2,
            getLastRow: () => sheet.getLastRow()
          };
          colorizeSheet_(viewSheetObj);
        }
        */
      } catch (e) {
        Logger.log(`シート「${sheetName}」の色付け処理中にエラー: ${e.message}`);
      }
    });
    logWithTimestamp("メインシートと工数シートの色付け処理が完了しました。");
  } catch (error) {
    Logger.log(`色付け処理でエラーが発生しました: ${error.stack}`);
  }
}


function colorizeSheet_(sheetObject) {
  // QUERYで生成されるビューはメインシートの書式を継承するため、この関数からは除外
  if (sheetObject.getSheet().getName().startsWith('ソートビュー')) {
      return;
  }
    
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

      if (sheetObject instanceof MainSheet) {
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