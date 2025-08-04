/**
 * Code.gs
 * イベントハンドラとカスタムメニューを管理する司令塔
 */

// =================================================================================
// === イベントハンドラ (司令塔) ===
// =================================================================================

function flowManager(e) {
  if (!e || !e.source || !e.range) { return; }
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const sheetName = sheet.getName();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ss.toast(`処理を開始します... (${sheetName}シート編集中)`, "情報", 3);

  let actionPerformed = false;
  
  // ===============================================================
  // === メインシートが編集された場合の処理 ===
  // ===============================================================
  if (sheetName === CONFIG.SHEETS.MAIN) {
    const mainIndices = getColumnIndices(sheet, MAIN_SHEET_HEADERS);
    const editedCol = range.getColumn();
    const editedRow = range.getRow();

    // 担当者、問い合わせ、または進捗が変更されたら工数シートに即時同期
    if (editedCol === mainIndices.TANTOUSHA || editedCol === mainIndices.TOIAWASE || editedCol === mainIndices.PROGRESS) {
       syncMainToInput();
    }
    
    // 進捗列が変更されたら更新日時を記録
    if (editedCol === mainIndices.PROGRESS && editedRow >= 2) {
      sheet.getRange(editedRow, mainIndices.UPDATE_TS).setValue(new Date());
    }

    actionPerformed = true;

  // ===============================================================
  // === 工数シートが編集された場合の処理 ===
  // ===============================================================
  } else if (sheetName.startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
    const inputIndices = getColumnIndices(sheet, INPUT_SHEET_HEADERS);

    // 進捗列が編集されたらタイムスタンプを記録
    if (range.getColumn() === inputIndices.PROGRESS && range.getRow() >= 3) {
      sheet.getRange(range.getRow(), inputIndices.TIMESTAMP).setValue(new Date());
    }
    syncProgressFromInputToMain(); // 工数シートの進捗をメインシートへ同期
    actionPerformed = true;
  }

  // ===============================================================
  // === 共通の後処理 ===
  // ===============================================================
  if (actionPerformed) {
    // 変更があった場合のみ実行する後続処理
    const mainSheet = ss.getSheetByName(CONFIG.SHEETS.MAIN);
    if(mainSheet){
      const mainIndices = getColumnIndices(mainSheet, MAIN_SHEET_HEADERS);
      // メインシートの値を基に全体を更新・整形
      rebuildInputSheetsFromMainOptimized();
      batchUpdateCompletionDates();
      updateMainSheetLaborTotal();
      checkAndHandleDuplicateMachineNumbers();
      
      // 色付け処理
      colorizeManagementNoByProgressInMainSheet();
      colorizeProgressColumnInMainSheet();
      colorizeManagementNoInInputSheets();
      colorizeTantoushaCellInInputSheets();
      colorizeToiawaseCellInInputSheets();
    }
    ss.toast("自動処理が完了しました。", "完了", 3);
  }
}

// =================================================================================
// === カスタムメニューと追加機能 ===
// =================================================================================

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('カスタムメニュー')
      .addItem('操作パネルを開く (工数シート月表示)', 'showControlSidebar')
      .addSeparator()
      .addItem('全機番の資料フォルダ作成', 'bulkCreateKibanFolders')
      .addItem('全機種シリーズのフォルダ作成', 'bulkCreateSeriesFolders')
      .addSeparator()
      .addItem('週次バックアップを作成', 'createWeeklyBackup')
      .addToUi();
}

function showControlSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('操作パネル');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getMonthsFromLaborSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(CONFIG.SHEETS.INPUT_PREFIX));
  const months = new Set();

  inputSheets.forEach(sheet => {
    const inputIndices = getColumnIndices(sheet, INPUT_SHEET_HEADERS);
    const laborStartCol = inputIndices.TOTAL_HOURS ? inputIndices.TOTAL_HOURS + 1 : -1;
    if(laborStartCol === -1) return;

    const lastCol = sheet.getLastColumn();
    if (lastCol < laborStartCol) return;
    
    const headerDates = sheet.getRange(1, laborStartCol, 1, lastCol - laborStartCol + 1).getValues()[0];
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
  const inputSheets = ss.getSheets().filter(sheet => sheet.getName().startsWith(CONFIG.SHEETS.INPUT_PREFIX));

  inputSheets.forEach(sheet => {
    const inputIndices = getColumnIndices(sheet, INPUT_SHEET_HEADERS);
    const laborStartCol = inputIndices.TOTAL_HOURS ? inputIndices.TOTAL_HOURS + 1 : -1;
    if(laborStartCol === -1) return;

    const lastCol = sheet.getLastColumn();
    if (lastCol < laborStartCol) return;
    
    sheet.showColumns(laborStartCol, lastCol - laborStartCol + 1);
    
    if (selectedMonth !== "all") {
      const [targetYear, targetMonth] = selectedMonth.split('-').map(Number);
      const headerDates = sheet.getRange(1, laborStartCol, 1, lastCol - laborStartCol + 1).getValues()[0];
      
      headerDates.forEach((date, i) => {
        if(date instanceof Date && !isNaN(date)){
          const d = new Date(date);
          if (!(d.getFullYear() === targetYear && d.getMonth() === targetMonth)) {
            sheet.hideColumns(laborStartCol + i);
          }
        }
      });
    }
  });
}