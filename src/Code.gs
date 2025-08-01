// =================================================================================
// === イベントハンドラ (司令塔) (最終レイアウト対応版) ===
// =================================================================================
function flowManager(e) {
  if (!e || !e.source || !e.range) { return; }
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const sheetName = sheet.getName();
  
  SpreadsheetApp.getActiveSpreadsheet().toast('処理を開始します... (' + sheetName + 'シート編集中)', "情報", 3);
  let actionPerformed = false;

  if (sheetName === KIBAN_MASTER_SHEET_NAME) {
    updateMainSheetFromMaster();
    rebuildInputSheetsFromMainOptimized();
    actionPerformed = true;
  } else if (sheetName === MAIN_SHEET_NAME) {
    // ★★★ ここからが修正箇所です ★★★
    const editedCol = range.getColumn();
    const editedRow = range.getRow();

    // もし進捗列(M列)が編集されたら、同じ行の更新日時(P列)を現在の時刻に更新する
    if (editedCol === MAIN_SHEET_PROGRESS_COL && editedRow >= 2) {
      sheet.getRange(editedRow, MAIN_SHEET_UPDATE_TS_COL).setValue(new Date());
    }
    // ★★★ ここまでが修正箇所です ★★★

    if (editedCol === MAIN_SHEET_TANTOUSHA_COL || editedCol === MAIN_SHEET_TOIAWASE_COL) {
      syncContactInfoFromMainToInputSheets();
    }
    actionPerformed = true;
  } else if (sheetName.startsWith(INPUT_SHEET_PREFIX)) {
    if (range.getColumn() === INPUT_SHEET_PROGRESS_COL && range.getRow() >= 3) {
      sheet.getRange(range.getRow(), INPUT_SHEET_TIMESTAMP_COL).setValue(new Date());
    }
    syncProgressFromInputToMain();
    updateMainSheetLaborTotal();
    actionPerformed = true;
  } else if (sheetName === '生産管理マスタ') {
    syncAssemblyStartDateFromProdMaster();
    actionPerformed = true;
  }

  if (actionPerformed) {
    // 全体的な更新処理をまとめる
    applyProgressPlaceholder();
    batchUpdateCompletionDates();
    syncProgressToMaster();
    markOrphanedRowsInInputSheets();
    
    checkAndHandleDuplicateMachineNumbers();
    syncProgressFromMainToInput();

    colorizeManagementNoByProgressInMainSheet();
    colorizeProgressColumnInMainSheet();
    colorizeManagementNoInInputSheets();
    colorizeTantoushaCellInInputSheets();
    colorizeToiawaseCellInInputSheets();
    SpreadsheetApp.getActiveSpreadsheet().toast("自動処理が完了しました。", "完了", 3);
  }
}

function syncContactInfoFromMainToInputSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!mainSheet) return;
  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(INPUT_SHEET_PREFIX));
  const mainLastRow = mainSheet.getLastRow();
  if (mainLastRow <= 1) return;
  
  const mainData = mainSheet.getRange(2, 1, mainLastRow - 1, Math.max(MAIN_SHEET_TANTOUSHA_COL, MAIN_SHEET_TOIAWASE_COL)).getValues();
  const mainInfoMap = new Map();
  mainData.forEach(row => {
    const mgmtNo = String(row[MAIN_SHEET_MGMT_NO_COL - 1]).trim();
    if (mgmtNo) {
      mainInfoMap.set(mgmtNo, {
        tantousha: row[MAIN_SHEET_TANTOUSHA_COL - 1],
        toiawase: row[MAIN_SHEET_TOIAWASE_COL - 1]
      });
    }
  });

  inputSheets.forEach(sheet => {
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    const range = sheet.getRange(3, 1, lastRow - 2, Math.max(INPUT_SHEET_TANTOU_COL, INPUT_SHEET_TOIAWASE_COL));
    const values = range.getValues();
    let updated = false;

    const newValues = values.map(row => {
      const mgmtNo = String(row[INPUT_SHEET_MGMT_NO_COL - 1]).trim();
      if (mainInfoMap.has(mgmtNo)) {
        const info = mainInfoMap.get(mgmtNo);
        if (row[INPUT_SHEET_TANTOU_COL - 1] !== info.tantousha) {
          row[INPUT_SHEET_TANTOU_COL - 1] = info.tantousha;
          updated = true;
        }
        if (row[INPUT_SHEET_TOIAWASE_COL - 1] !== info.toiawase) {
          row[INPUT_SHEET_TOIAWASE_COL - 1] = info.toiawase;
          updated = true;
        }
      }
      return row; // このままだと全行を更新してしまうので修正
    });

    if (updated) {
       const tantouValues = values.map(row => [row[INPUT_SHEET_TANTOU_COL - 1]]);
       const toiawaseValues = values.map(row => [row[INPUT_SHEET_TOIAWASE_COL - 1]]);
       sheet.getRange(3, INPUT_SHEET_TANTOU_COL, tantouValues.length, 1).setValues(tantouValues);
       sheet.getRange(3, INPUT_SHEET_TOIAWASE_COL, toiawaseValues.length, 1).setValues(toiawaseValues);
    }
  });
}

// =================================================================================
// === カスタムメニューと追加機能 ===
// =================================================================================
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('カスタムメニュー')
      .addItem('操作パネルを開く (工数シート月表示)', 'showControlSidebar')
      .addSeparator()
      .addItem('全機番の資料フォルダ作成 (機番マスタH列)', 'bulkCreateKibanFolders')
      .addItem('全機種シリーズのフォルダ作成 (機番マスタI列)', 'bulkCreateSeriesFolders')
      .addToUi();
}

function showControlSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('操作パネル');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getMonthsFromLaborSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(INPUT_SHEET_PREFIX));
  const months = new Set();
  
  inputSheets.forEach(sheet => {
    const lastCol = sheet.getLastColumn();
    if (lastCol < INPUT_SHEET_LABOR_START_COL) return;
    const headerDates = sheet.getRange(1, INPUT_SHEET_LABOR_START_COL, 1, lastCol - INPUT_SHEET_LABOR_START_COL + 1).getValues()[0];
    headerDates.forEach(date => {
      if (date instanceof Date && !isNaN(date)) {
        months.add(date.getFullYear() + '-' + date.getMonth());
      }
    });
  });
  
  return Array.from(months).map(m => {
    const [year, month] = m.split('-');
    return { text: `${year}年${parseInt(month, 10) + 1}月`, value: m };
  }).sort((a, b) => new Date(a.value.split('-')[0], a.value.split('-')[1]) - new Date(b.value.split('-')[0], b.value.split('-')[1]));
}

function filterLaborSheetColumnsByMonth(selectedMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(INPUT_SHEET_PREFIX));
  
  inputSheets.forEach(sheet => {
    const lastCol = sheet.getLastColumn();
    if (lastCol < INPUT_SHEET_LABOR_START_COL) return;
    sheet.showColumns(INPUT_SHEET_LABOR_START_COL, lastCol - INPUT_SHEET_LABOR_START_COL + 1);
    
    if (selectedMonth !== "all") {
      const [targetYear, targetMonth] = selectedMonth.split('-').map(Number);
      const headerDates = sheet.getRange(1, INPUT_SHEET_LABOR_START_COL, 1, lastCol - INPUT_SHEET_LABOR_START_COL + 1).getValues()[0];
      headerDates.forEach((date, i) => {
        if(date instanceof Date && !isNaN(date)){
          const d = new Date(date);
          if (!(d.getFullYear() === targetYear && d.getMonth() === targetMonth)) {
            sheet.hideColumns(INPUT_SHEET_LABOR_START_COL + i);
          }
        }
      });
    }
  });
}