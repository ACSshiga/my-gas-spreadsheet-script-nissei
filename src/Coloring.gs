/**
 * Coloring.gs
 * シートの自動色付けに関する機能を管理します。
 * データの状態を視覚的に分かりやすくします。
 */

// =================================================================================
// === 色付け処理（メイン） ===
// =================================================================================

function colorizeAllSheets() {
  try {
    const mainSheet = new MainSheet();
    colorizeSheet_(mainSheet);

    const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    allSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName.startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
        try {
          const tantoushaName = sheetName.replace(CONFIG.SHEETS.INPUT_PREFIX, '');
          const inputSheet = new InputSheet(tantoushaName);
          colorizeSheet_(inputSheet);
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

/**
 * ★★★ 最終版 ★★★
 * 指定されたシートオブジェクトの各列をルールに基づき色付けします。
 * 数式を保護しながら、重複行（同じ機番＋作業区分）の検出と処理を行います。
 */
function colorizeSheet_(sheetObject) {
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
      // 既存の色をリセット
      for (let j = 0; j < lastCol; j++) {
          if(!formulas[i][j]) backgroundColors[i][j] = CONFIG.COLORS.DEFAULT_BACKGROUND;
      }

      if (progressCol) {
        const progressColor = getColor(PROGRESS_COLORS, safeTrim(row[progressCol - 1]));
        backgroundColors[i][progressCol - 1] = progressColor;
        if (mgmtNoCol) backgroundColors[i][mgmtNoCol - 1] = progressColor;
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
}

/**
 * 工数シートの土日・祝日の日付列に背景色を設定します。
 */
function colorizeHolidayColumns_(inputSheetObject) {
  const sheet = inputSheetObject.getSheet();
  const lastCol = sheet.getLastColumn();
  const dateColumnStart = Object.keys(INPUT_SHEET_HEADERS).length + 1;
  if (lastCol < dateColumnStart) return;
  const year = new Date().getFullYear();
  const holidays = getJapaneseHolidays(year);
  const nextYearHolidays = getJapaneseHolidays(year + 1);
  nextYearHolidays.forEach(h => holidays.add(h));
  const headerRange = sheet.getRange(1, dateColumnStart, 1, lastCol - dateColumnStart + 1);
  const headerDates = headerRange.getValues()[0];
  for (let i = 0; i < headerDates.length; i++) {
    const currentCol = dateColumnStart + i;
    const date = headerDates[i];
    const color = (isValidDate(date) && isHoliday(date, holidays)) ? CONFIG.COLORS.WEEKEND_HOLIDAY : CONFIG.COLORS.DEFAULT_BACKGROUND;
    const colBackgrounds = Array(sheet.getMaxRows()).fill([color]);
    sheet.getRange(1, currentCol, sheet.getMaxRows()).setBackgrounds(colBackgrounds);
  }
}