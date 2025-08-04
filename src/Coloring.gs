// =================================================================================
// === 色付け・視覚的補助 (最終レイアウト対応版) ===
// =================================================================================
function colorizeLaborInputColumnsOnHolidaysAndWeekends() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const holidayAndWeekendColor = DUPLICATE_HIGHLIGHT_COLOR;
  const defaultBackgroundColor = DEFAULT_CELL_BACKGROUND_COLOR;
  const calendar = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
  if (!calendar) return;

  const today = new Date();
  const events = calendar.getEvents(new Date(today.getFullYear() - 1, 0, 1), new Date(today.getFullYear() + 1, 11, 31));
  const holidayDates = new Set(events.map(e => formatDateForComparison(e.getStartTime())));

  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(INPUT_SHEET_PREFIX));
  inputSheets.forEach(sheet => {
    const lastCol = sheet.getLastColumn();
    if (lastCol < INPUT_SHEET_LABOR_START_COL) return;

    const headerDates = sheet.getRange(1, INPUT_SHEET_LABOR_START_COL, 1, lastCol - INPUT_SHEET_LABOR_START_COL + 1).getValues()[0];
    const rangeToColor = sheet.getRange(1, INPUT_SHEET_LABOR_START_COL, sheet.getMaxRows(), lastCol - INPUT_SHEET_LABOR_START_COL + 1);
    const backgrounds = rangeToColor.getBackgrounds();
    
    headerDates.forEach((date, i) => {
      if(date){
        const day = new Date(date).getDay();
        const isHoliday = holidayDates.has(formatDateForComparison(new Date(date))) || day === 0 || day === 6;
        if (isHoliday) {
          for (let j = 0; j < backgrounds.length; j++) {
            backgrounds[j][i] = holidayAndWeekendColor;
          }
        }
      }
    });
    rangeToColor.setBackgrounds(backgrounds);
  });
}

function colorizeManagementNoByProgressInMainSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  
  const colorRules = { "未着手": "#ffcfc9", "配置済": "#d4edbc", "ACS済": "#bfe1f6", "日精済": "#ffe5a0", "係り中": "#e6cff2", "機番重複": DUPLICATE_HIGHLIGHT_COLOR, "保留": "#c6dbe1" };
  const range = sheet.getRange(2, 1, lastRow - 1, MAIN_SHEET_PROGRESS_COL);
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();

  const newBgs = values.map((row, i) => {
    const progress = String(row[MAIN_SHEET_PROGRESS_COL - 1]).trim();
    // 機番重複の背景色は checkAndHandleDuplicateMachineNumbers で設定されるため、ここでは上書きしない
    // ただし、重複が解消された場合は、新しい進捗の色を適用する必要がある
    const kiban = String(row[MAIN_SHEET_KIBAN_COL-1]).trim();
    const currentBg = backgrounds[i][MAIN_SHEET_MGMT_NO_COL - 1];

    if (progress === '機番重複') {
      return [DUPLICATE_HIGHLIGHT_COLOR];
    }
    return [colorRules[progress] || DEFAULT_CELL_BACKGROUND_COLOR];
  });
  
  sheet.getRange(2, MAIN_SHEET_MGMT_NO_COL, newBgs.length, 1).setBackgrounds(newBgs);
}

function colorizeProgressColumnInMainSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  
  const colorRules = { "未着手": "#ffcfc9", "配置済": "#d4edbc", "ACS済": "#bfe1f6", "日精済": "#ffe5a0", "係り中": "#e6cff2", "機番重複": DUPLICATE_HIGHLIGHT_COLOR, "保留": "#c6dbe1" };
  const range = sheet.getRange(2, MAIN_SHEET_PROGRESS_COL, lastRow - 1, 1);
  const values = range.getValues();
  
  const newBgs = values.map(row => {
    const progress = String(row[0]).trim();
    return [colorRules[progress] || DEFAULT_CELL_BACKGROUND_COLOR];
  });
  
  range.setBackgrounds(newBgs);
}

function colorizeManagementNoInInputSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(INPUT_SHEET_PREFIX));
  const colorRules = { "未着手": "#ffcfc9", "配置済": "#d4edbc", "ACS済": "#bfe1f6", "日精済": "#ffe5a0", "係り中": "#e6cff2", "機番重複": DUPLICATE_HIGHLIGHT_COLOR, "保留": "#c6dbe1" };

  inputSheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    const range = sheet.getRange(3, 1, lastRow - 2, INPUT_SHEET_PROGRESS_COL);
    const values = range.getValues();
    const backgrounds = range.getBackgrounds();

    const newBgs = values.map((row, i) => {
      const progress = String(row[INPUT_SHEET_PROGRESS_COL - 1]).trim();
      const currentBg = backgrounds[i][0];
      if (progress === '機番重複') {
          return [DUPLICATE_HIGHLIGHT_COLOR];
      }
      return [colorRules[progress] || DEFAULT_CELL_BACKGROUND_COLOR];
    });

    sheet.getRange(3, INPUT_SHEET_MGMT_NO_COL, newBgs.length, 1).setBackgrounds(newBgs);
  });
}

function colorizeTantoushaCellInInputSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(INPUT_SHEET_PREFIX));
  const colorRules = { "志賀": "#ffcfc9", "遠藤": "#d4edbc", "小板橋": "#bfe1f6" };

  inputSheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    const range = sheet.getRange(3, 1, lastRow - 2, INPUT_SHEET_TANTOU_COL);
    const values = range.getValues();
    const backgrounds = range.getBackgrounds();
    
    const newBgs = values.map((row, i) => {
      const tantousha = String(row[INPUT_SHEET_TANTOU_COL - 1]).trim();
      return [colorRules[tantousha] || DEFAULT_CELL_BACKGROUND_COLOR];
    });
    
    sheet.getRange(3, INPUT_SHEET_TANTOU_COL, newBgs.length, 1).setBackgrounds(newBgs);
  });
}

function colorizeToiawaseCellInInputSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(INPUT_SHEET_PREFIX));
  const colorRules = { "問合済": "#ffcfc9", "回答済": "#bfe1f6" };
  
  inputSheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    const range = sheet.getRange(3, 1, lastRow - 2, INPUT_SHEET_TOIAWASE_COL);
    const values = range.getValues();
    
    const newBgs = values.map((row, i) => {
      const toiawase = String(row[INPUT_SHEET_TOIAWASE_COL - 1]).trim();
      return [colorRules[toiawase] || DEFAULT_CELL_BACKGROUND_COLOR];
    });
    
    sheet.getRange(3, INPUT_SHEET_TOIAWASE_COL, newBgs.length, 1).setBackgrounds(newBgs);
  });
}/**
 * Coloring.gs
 * シートの自動色付けに関する機能を管理
 */

// =================================================================================
// === 色付け処理（メイン） ===
// =================================================================================

/**
 * すべてのシートの色付けをまとめて実行するメイン関数。
 */
function colorizeAllSheets() {
  const mainSheet = new MainSheet();
  const tantoushaList = mainSheet.getTantoushaList();

  // メインシートの色付け
  colorizeSheet_(mainSheet);

  // 各工数シートの色付け
  tantoushaList.forEach(tantousha => {
    try {
      const inputSheet = new InputSheet(tantousha);
      colorizeSheet_(inputSheet);
    } catch (e) {
      // シートが存在しない場合はスキップ
    }
  });

  // 工数シートの日付列の色付け
  colorizeHolidayColumns_();
  logWithTimestamp("全シートの色付け処理が完了しました。");
}

/**
 * 指定されたシートオブジェクトの各列をルールに基づいて色付けする内部関数。
 * @param {MainSheet | InputSheet} sheetObject - 色付け対象のシートオブジェクト
 */
function colorizeSheet_(sheetObject) {
  const sheet = sheetObject.getSheet();
  const indices = sheetObject.indices;
  const lastRow = sheetObject.getLastRow();
  const startRow = (sheetObject instanceof MainSheet) ? 2 : 2; // データ開始行

  if (lastRow < startRow) return;

  const range = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheetObject.getLastColumn());
  const values = range.getValues();

  // 色情報を格納する2次元配列を準備
  const backgroundColors = range.getBackgrounds();

  values.forEach((row, i) => {
    // 各列の値を取得
    const progressPanel = safeTrim(row[indices.PROGRESS_PANEL - 1]);
    const progressWire = safeTrim(row[indices.PROGRESS_WIRE - 1]);
    const tantousha = safeTrim(row[indices.TANTOUSHA - 1]);
    const toiawase = safeTrim(row[indices.TOIAWASE - 1]);

    // 色を決定
    const mgmtNoColor = (progressPanel === '完了' && progressWire === '完了') 
      ? getColor(PROGRESS_COLORS, '完了') 
      : (progressPanel === '未着手' && progressWire === '未着手') 
        ? getColor(PROGRESS_COLORS, '未着手') 
        : getColor(PROGRESS_COLORS, '仕掛中');

    // 配列の対応する位置に色情報をセット
    backgroundColors[i][indices.MGMT_NO - 1] = mgmtNoColor;
    backgroundColors[i][indices.PROGRESS_PANEL - 1] = getColor(PROGRESS_COLORS, progressPanel);
    backgroundColors[i][indices.PROGRESS_WIRE - 1] = getColor(PROGRESS_COLORS, progressWire);
    backgroundColors[i][indices.TANTOUSHA - 1] = getColor(TANTOUSHA_COLORS, tantousha);
    backgroundColors[i][indices.TOIAWASE - 1] = getColor(TOIAWASE_COLORS, toiawase);
  });

  // 範囲に背景色を一度に設定
  range.setBackgrounds(backgroundColors);
}

/**
 * 工数シートの土日・祝日の日付列に背景色を設定します。
 */
function colorizeHolidayColumns_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = new MainSheet();
  const tantoushaList = mainSheet.getTantoushaList();
  
  const holidays = getJapaneseHolidays(new Date(new Date().getFullYear(), 0, 1), new Date(new Date().getFullYear() + 1, 11, 31));

  tantoushaList.forEach(tantousha => {
    try {
      const inputSheet = new InputSheet(tantousha);
      const sheet = inputSheet.getSheet();
      const lastCol = inputSheet.getLastColumn();
      
      const dateColumnStart = Object.keys(INPUT_SHEET_HEADERS).length + 1;
      if (lastCol < dateColumnStart) return;

      const headerRange = sheet.getRange(1, dateColumnStart, 1, lastCol - dateColumnStart + 1);
      const headerDates = headerRange.getValues()[0];

      headerDates.forEach((date, i) => {
        const currentCol = dateColumnStart + i;
        if (isValidDate(date) && isHoliday(date, holidays)) {
          sheet.getRange(1, currentCol, sheet.getMaxRows()).setBackground(CONFIG.COLORS.WEEKEND_HOLIDAY);
        }
      });
    } catch (e) {
      // シートが存在しない場合はスキップ
    }
  });
}