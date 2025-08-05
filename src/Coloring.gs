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
 * ★★★ ロジック全体を修正 ★★★
 * 指定されたシートオブジェクトの各列をルールに基づき色付けします。
 * 重複行（同じ機番＋作業区分）を検出し、行を灰色にし、進捗名を変更し、担当者セルをロックします。
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
  
  const values = fullRange.getValues();
  const backgroundColors = fullRange.getBackgrounds();

  // --- 列のインデックスを取得 ---
  const mgmtNoCol = indices.MGMT_NO;
  const progressCol = indices.PROGRESS;
  const tantoushaCol = indices.TANTOUSHA;
  const toiawaseCol = indices.TOIAWASE;
  const kibanCol = indices.KIBAN;
  const sagyouKubunCol = indices.SAGYOU_KUBUN;

  const DUPLICATE_COLOR = '#cccccc'; // 灰色

  const uniqueKeys = new Set();
  const restrictedRanges = []; // 担当者入力を禁止するセルのリスト
  const normalRanges = []; // 担当者入力を許可するセルのリスト

  values.forEach((row, i) => {
    let isDuplicate = false;
    // --- 重複チェック ---
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
    
    // --- 色と値の設定 ---
    if (isDuplicate) {
      //【重複時の処理】
      // 1. 行全体を灰色にする
      for (let j = 0; j < lastCol; j++) {
        backgroundColors[i][j] = DUPLICATE_COLOR;
      }
      // 2. 進捗の値を「機番重複」に変更
      if (progressCol) values[i][progressCol - 1] = "機番重複";
      // 3. 担当者の値をクリア
      if (tantoushaCol) values[i][tantoushaCol - 1] = "";
      
      // 4. メインシートの場合、担当者セルをロック対象に追加
      if (sheetObject instanceof MainSheet && tantoushaCol) {
        restrictedRanges.push(sheet.getRange(startRow + i, tantoushaCol));
      }

    } else {
      //【通常時の処理】
      // 1. 基本的な色付け
      if (progressCol) {
        const progressColor = getColor(PROGRESS_COLORS, safeTrim(row[progressCol - 1]));
        backgroundColors[i][progressCol - 1] = progressColor;
        if (mgmtNoCol) backgroundColors[i][mgmtNoCol - 1] = progressColor;
      }
      if (sheetObject instanceof MainSheet || sheet.getName().startsWith('View_')) {
        if (tantoushaCol) backgroundColors[i][tantoushaCol - 1] = getColor(TANTOUSHA_COLORS, safeTrim(row[tantoushaCol - 1]));
        if (toiawaseCol) backgroundColors[i][toiawaseCol - 1] = getColor(TOIAWASE_COLORS, safeTrim(row[toiawaseCol - 1]));
      }
      // 2. メインシートの場合、担当者セルをロック解除対象に追加
      if (sheetObject instanceof MainSheet && tantoushaCol) {
        normalRanges.push(sheet.getRange(startRow + i, tantoushaCol));
      }
    }
  });

  // --- 最後に色と値をまとめてシートに適用 ---
  fullRange.setBackgrounds(backgroundColors);
  fullRange.setValues(values);

  // --- 担当者列の入力制限を適用 ---
  if (sheetObject instanceof MainSheet) {
    // 1. ロックするセルの設定
    if (restrictedRanges.length > 0) {
      const restrictedRule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(sheet.getRange('A1:A1'), false) // ダミーの参照範囲
        .setAllowInvalid(false)
        .setHelpText('この行は機番が重複しているため、担当者は設定できません。')
        .build();
      restrictedRanges.forEach(range => range.setDataValidation(restrictedRule));
    }
    // 2. ロックを解除するセルの設定
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
}

/**
 * 工数シートの土日・祝日の日付列に背景色を設定します。
 * (この関数は現在、パフォーマンス上の理由でcolorizeAllSheetsからは呼び出されていません)
 */
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
  const backgrounds = [];
  for (let i = 0; i < headerDates.length; i++) {
    const currentCol = dateColumnStart + i;
    const date = headerDates[i];
    const color = (isValidDate(date) && isHoliday(date, holidays)) ? CONFIG.COLORS.WEEKEND_HOLIDAY : CONFIG.COLORS.DEFAULT_BACKGROUND;
    const colBackgrounds = Array(sheet.getMaxRows()).fill([color]);
    sheet.getRange(1, currentCol, sheet.getMaxRows()).setBackgrounds(colBackgrounds);
  }
}