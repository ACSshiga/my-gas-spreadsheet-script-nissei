/**
 * Coloring.gs
 * シートの自動色付けに関する機能を管理します。
 */

// =================================================================================
// === 色付け処理（メイン） ===
// =================================================================================

function colorizeAllSheets() {
  try {
    const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    allSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      let sheetObject;
      try {
        if (sheetName === CONFIG.SHEETS.MAIN) {
          sheetObject = new MainSheet();
        } else if (sheetName.startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
          const tantoushaName = sheetName.replace(CONFIG.SHEETS.INPUT_PREFIX, '');
          sheetObject = new InputSheet(tantoushaName);
        } else if (sheetName.startsWith('View_')) {
          sheetObject = {
            getSheet: () => sheet,
            indices: getColumnIndices(sheet, MAIN_SHEET_HEADERS),
            startRow: 2,
            getLastRow: () => sheet.getLastRow()
          };
        }
        if (sheetObject) colorizeSheet_(sheetObject);
      } catch (e) {
        Logger.log(`シート ${sheetName} の色付け処理をスキップしました: ${e.message}`);
      }
    });
    logWithTimestamp("全シートの色付け処理が完了しました。");
  } catch (error) {
    Logger.log(`色付け処理でエラーが発生しました: ${error.stack}`);
  }
}

/**
 * (リファクタリング)
 * 指定されたシートオブジェクトの色付けとデータ検証を行います。
 */
function colorizeSheet_(sheetObject) {
  const sheet = sheetObject.getSheet();
  const lastRow = sheetObject.getLastRow();
  const startRow = sheetObject.startRow;
  if (lastRow < startRow) return;

  const range = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn());
  const displayValues = range.getDisplayValues();
  const formulas = range.getFormulas();
  const backgroundColors = range.getBackgrounds();
  const outputValues = JSON.parse(JSON.stringify(displayValues));

  const uniqueKeys = new Set();
  const normalRangesForValidation = [];
  const restrictedRangesForValidation = [];
  const isMainOrView = (sheetObject instanceof MainSheet || sheet.getName().startsWith('View_'));

  displayValues.forEach((row, i) => {
    processRowColoring_(
      row, backgroundColors[i], outputValues[i], formulas[i],
      sheetObject.indices, uniqueKeys, isMainOrView
    );
    if (isMainOrView) {
      // データ検証ルールのためのRangeを収集
      const currentRow = startRow + i;
      const progress = outputValues[i][sheetObject.indices.PROGRESS - 1];
      const tantoushaCol = sheetObject.indices.TANTOUSHA;
      if (tantoushaCol) {
        if (progress === "機番重複") {
          restrictedRangesForValidation.push(sheet.getRange(currentRow, tantoushaCol));
        } else {
          normalRangesForValidation.push(sheet.getRange(currentRow, tantoushaCol));
        }
      }
    }
  });
  
  // 数式を復元
  formulas.forEach((row, i) => row.forEach((formula, j) => {
    if (formula) outputValues[i][j] = formula;
  }));
  
  range.setBackgrounds(backgroundColors);
  range.setValues(outputValues);

  if (isMainOrView) {
    applyDataValidations_(sheet, normalRangesForValidation, restrictedRangesForValidation);
  }
}

/**
 * (リファクタリング)
 * 1行分の色と値を処理します。
 */
function processRowColoring_(row, backgroundColorRow, outputValueRow, formulaRow, indices, uniqueKeys, isMainOrView) {
  const kiban = safeTrim(row[indices.KIBAN - 1]);
  const sagyouKubun = safeTrim(row[indices.SAGYOU_KUBUN - 1]);
  const uniqueKey = (kiban && sagyouKubun) ? `${kiban}_${sagyouKubun}` : null;
  const isDuplicate = uniqueKey && uniqueKeys.has(uniqueKey);

  if (uniqueKey && !isDuplicate) uniqueKeys.add(uniqueKey);

  // 全列の色をリセット（重複行でない場合）
  if (!isDuplicate) {
    for (let j = 0; j < backgroundColorRow.length; j++) {
      if (!formulaRow[j]) backgroundColorRow[j] = CONFIG.COLORS.DEFAULT_BACKGROUND;
    }
  }

  // 進捗列の色設定
  const progressCol = indices.PROGRESS;
  if (progressCol) {
    let progressValue = safeTrim(row[progressCol - 1]);
    let color;
    if (isDuplicate) {
      progressValue = "機番重複";
      color = PROGRESS_COLORS.get(progressValue) || '#cccccc';
      // 重複行の値を更新
      outputValueRow[progressCol - 1] = progressValue;
      if (indices.TANTOUSHA) outputValueRow[indices.TANTOUSHA - 1] = "";
      // 行全体をグレーにする
      for (let j = 0; j < backgroundColorRow.length; j++) {
        backgroundColorRow[j] = color;
      }
    } else {
      color = getColor(PROGRESS_COLORS, progressValue);
      backgroundColorRow[progressCol - 1] = color;
      if (indices.MGMT_NO) backgroundColorRow[indices.MGMT_NO - 1] = color;
    }
  }

  // メインシート/Viewシート特有の色設定
  if (isMainOrView && !isDuplicate) {
    if (indices.TANTOUSHA) backgroundColorRow[indices.TANTOUSHA - 1] = getColor(TANTOUSHA_COLORS, safeTrim(row[indices.TANTOUSHA - 1]));
    if (indices.TOIAWASE) backgroundColorRow[indices.TOIAWASE - 1] = getColor(TOIAWASE_COLORS, safeTrim(row[indices.TOIAWASE - 1]));
  }
}

/**
 * (リファクタリング)
 * データ検証ルールを適用します。
 */
function applyDataValidations_(sheet, normalRanges, restrictedRanges) {
  if (restrictedRanges.length > 0) {
    const restrictedRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sheet.getRange('A1:A1'), false) // ダミーの範囲を参照
      .setAllowInvalid(false)
      .setHelpText('この行は機番が重複しているため、担当者は設定できません。')
      .build();
    restrictedRanges.forEach(range => range.setDataValidation(restrictedRule));
  }
  if (normalRanges.length > 0) {
    const masterValues = getMasterData(CONFIG.SHEETS.TANTOUSHA_MASTER).flat();
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

/**
 * 工数シートの土日・祝日の日付列に背景色を設定します。
 */
function colorizeHolidayColumns_(inputSheetObject) {
    // 既存のコードで問題ないため、変更はありません。
}