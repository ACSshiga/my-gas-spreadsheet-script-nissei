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
}