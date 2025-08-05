/**
 * Coloring.gs
 * シートの自動色付けに関する機能を管理します。
 * データの状態を視覚的に分かりやすくします。
 */

// =================================================================================
// === 色付け処理（メイン） ===
// =================================================================================

/**
 * すべてのシートの色付けをまとめて実行するメイン関数です。
 * onEditの最後に呼び出すことで、常に最新の状態を保ちます。
 */
function colorizeAllSheets() {
  try {
    const mainSheet = new MainSheet();
    colorizeSheet_(mainSheet);

    // 工数シートの色付けは、存在するシートのみ処理
    const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    allSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName.startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
        try {
          const tantoushaName = sheetName.replace(CONFIG.SHEETS.INPUT_PREFIX, '');
          const inputSheet = new InputSheet(tantoushaName);
          colorizeSheet_(inputSheet);
          // 土日祝日の色付けは一旦スキップ（重い処理のため）
          // colorizeHolidayColumns_(inputSheet);
        } catch (e) {
          // エラーは無視
        }
      } else if (sheetName.startsWith('View_')) {
        try {
          // Viewシート用の簡易オブジェクト
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
 * 指定されたシートオブジェクトの各列をルールに基づいて色付けする内部関数。
 */
function colorizeSheet_(sheetObject) {
  const sheet = sheetObject.getSheet();
  const indices = sheetObject.indices;
  const lastRow = sheetObject.getLastRow();
  const startRow = sheetObject.startRow;

  if (lastRow < startRow) return;

  // 必要な列のみ取得して処理を軽量化
  const mgmtNoCol = indices.MGMT_NO || 0;
  const progressCol = indices.PROGRESS || 0;
  const tantoushaCol = indices.TANTOUSHA || 0;
  const toiawaseCol = indices.TOIAWASE || 0;

  if (!mgmtNoCol || !progressCol) return;

  const dataRows = lastRow - startRow + 1;
  
  // バッチ処理で色を設定
  const isMainOrView = sheetObject instanceof MainSheet || sheet.getName().startsWith('View_');
  
  // 進捗の色設定
  if (progressCol > 0) {
    const progressRange = sheet.getRange(startRow, progressCol, dataRows, 1);
    const progressValues = progressRange.getValues();
    const progressColors = progressValues.map(row => [getColor(PROGRESS_COLORS, safeTrim(row[0]))]);
    progressRange.setBackgrounds(progressColors);
    
    // 管理No列も同じ色に
    sheet.getRange(startRow, mgmtNoCol, dataRows, 1).setBackgrounds(progressColors);
  }
  
  // メインシートとViewシートのみ担当者と問い合わせの色付け
  if (isMainOrView) {
    if (tantoushaCol > 0) {
      const tantoushaRange = sheet.getRange(startRow, tantoushaCol, dataRows, 1);
      const tantoushaValues = tantoushaRange.getValues();
      const tantoushaColors = tantoushaValues.map(row => [getColor(TANTOUSHA_COLORS, safeTrim(row[0]))]);
      tantoushaRange.setBackgrounds(tantoushaColors);
    }
    
    if (toiawaseCol > 0) {
      const toiawaseRange = sheet.getRange(startRow, toiawaseCol, dataRows, 1);
      const toiawaseValues = toiawaseRange.getValues();
      const toiawaseColors = toiawaseValues.map(row => [getColor(TOIAWASE_COLORS, safeTrim(row[0]))]);
      toiawaseRange.setBackgrounds(toiawaseColors);
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
  // 来年の祝日も念のため取得
  const nextYearHolidays = getJapaneseHolidays(year + 1);
  nextYearHolidays.forEach(h => holidays.add(h));
  
  const headerRange = sheet.getRange(1, dateColumnStart, 1, lastCol - dateColumnStart + 1);
  const headerDates = headerRange.getValues()[0];
  const backgrounds = [];

  // 各日付列の背景色配列を作成
  for (let i = 0; i < headerDates.length; i++) {
    const currentCol = dateColumnStart + i;
    const date = headerDates[i];
    const color = (isValidDate(date) && isHoliday(date, holidays)) ? CONFIG.COLORS.WEEKEND_HOLIDAY : CONFIG.COLORS.DEFAULT_BACKGROUND;
    const colBackgrounds = Array(sheet.getMaxRows()).fill([color]);
    // 2次元配列にするため、1列ずつ設定
    sheet.getRange(1, currentCol, sheet.getMaxRows()).setBackgrounds(colBackgrounds);
  }
}