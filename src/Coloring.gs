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
 * ★★★ 調査用コード ★★★
 * 詳細なログを出力して、重複検出の動作を確認します。
 */
function colorizeSheet_(sheetObject) {
  const sheet = sheetObject.getSheet();
  const indices = sheetObject.indices;
  const lastRow = sheetObject.getLastRow();
  const startRow = sheetObject.startRow;

  if (lastRow < startRow) return;

  Logger.log(`--- 色付け処理開始: ${sheet.getName()} ---`);

  const dataRows = lastRow - startRow + 1;
  const lastCol = sheet.getLastColumn();
  const fullRange = sheet.getRange(startRow, 1, dataRows, lastCol);
  
  const values = fullRange.getValues();
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
        Logger.log(`行 ${startRow + i}: チェック中のキー = "${uniqueKey}"`); // キーをログに出力

        if (uniqueKeys.has(uniqueKey)) {
          isDuplicate = true;
          Logger.log(`行 ${startRow + i}: ★★★ 重複を検出しました ★★★`); // 重複検出をログに出力
        } else {
          uniqueKeys.add(uniqueKey);
        }
      }
    }
    
    if (isDuplicate) {
      for (let j = 0; j < lastCol; j++) {
        backgroundColors[i][j] = DUPLICATE_COLOR;
      }
      if (progressCol) values[i][progressCol - 1] = "機番重複";
      if (tantoushaCol) values[i][tantoushaCol - 1] = "";
      if (sheetObject instanceof MainSheet && tantoushaCol) {
        restrictedRanges.push(sheet.getRange(startRow + i, tantoushaCol));
      }
    } else {
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

  fullRange.setBackgrounds(backgroundColors);
  fullRange.setValues(values);

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
      }
    }
  }
  Logger.log(`--- 色付け処理終了: ${sheet.getName()} ---`);
}

function colorizeHolidayColumns_(inputSheetObject) {
  // ... (この関数の内容は変更ありません)
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