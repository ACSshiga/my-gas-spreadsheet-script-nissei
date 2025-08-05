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

    const tantoushaList = mainSheet.getTantoushaList();
    tantoushaList.forEach(tantousha => {
      try {
        const inputSheet = new InputSheet(tantousha.name);
        colorizeSheet_(inputSheet);
        colorizeHolidayColumns_(inputSheet);
      } catch (e) {
        // シートが存在しない場合はスキップ
      }
    });
    
    // Viewシートの色付けを追加
    const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    allSheets.forEach(sheet => {
      if (sheet.getName().startsWith('View_')) {
        try {
          // Viewシート用の一時的なMainSheetオブジェクトを作成
          const viewSheetObj = {
            getSheet: () => sheet,
            indices: getColumnIndices(sheet, MAIN_SHEET_HEADERS),
            startRow: 2,
            getLastRow: () => sheet.getLastRow()
          };
          colorizeSheet_(viewSheetObj);
        } catch (e) {
          Logger.log(`Viewシート ${sheet.getName()} の色付けエラー: ${e.message}`);
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

  const range = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn());
  const values = range.getValues();
  // 現在の背景色を取得して、必要な部分だけ変更する
  const backgroundColors = range.getBackgrounds();

  values.forEach((row, i) => {
    // メインシートの場合のみ、担当者と問い合わせの色付けを行う
    if (sheetObject instanceof MainSheet) {
      const tantousha = safeTrim(row[indices.TANTOUSHA - 1]);
      const toiawase = safeTrim(row[indices.TOIAWASE - 1]);
      backgroundColors[i][indices.TANTOUSHA - 1] = getColor(TANTOUSHA_COLORS, tantousha);
      backgroundColors[i][indices.TOIAWASE - 1] = getColor(TOIAWASE_COLORS, toiawase);
    }
    
    // 両シート共通の色付け
    const progress = safeTrim(row[indices.PROGRESS - 1]);
    const progressColor = getColor(PROGRESS_COLORS, progress);
    backgroundColors[i][indices.MGMT_NO - 1] = progressColor; // 管理No.列
    backgroundColors[i][indices.PROGRESS - 1] = progressColor; // 進捗列
  });

  // 色情報をまとめて一度に設定
  range.setBackgrounds(backgroundColors);
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