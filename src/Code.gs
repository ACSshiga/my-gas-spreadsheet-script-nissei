/**
 * Code.gs
 * イベントハンドラとカスタムメニューを管理する司令塔。
 * スプレッドシートの操作をトリガーに、各機能を呼び出します。
 */

// =================================================================================
// === イベントハンドラ (スプレッドシート操作時に自動実行) ===
// =================================================================================

/**
 * スプレッドシートが開かれたときに実行される関数です。
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('カスタムメニュー')
    .addItem('自分の担当案件のみ表示', 'applyMyTasksFilter')
    .addItem('すべてのフィルタを解除', 'clearAllFilters')
    .addSeparator()
    .addItem('サイドバーを開く (フィルタ)', 'showFilterSidebar')
    .addSeparator()
    .addItem('請求シートを更新', 'showBillingSidebar') // 請求用サイドバーを開くように変更
    .addSeparator()
    .addItem('全機番の資料フォルダ作成', 'bulkCreateKibanFolders')
    .addItem('週次バックアップを作成', 'createWeeklyBackup')
    .addToUi();

  // 前回のフィルタ状態を復元
  restoreLastFilter();
  // 全マスタシートからデータ入力規則を更新
  setupAllDataValidations();
}

/**
 * スプレッドシートが編集されたときに実行される関数です。
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - イベントオブジェクト
 */
function onEdit(e) {
  if (!e || !e.source || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const ss = e.source;

  try {
    // メインシートが編集された場合
    if (sheetName === CONFIG.SHEETS.MAIN) {
      ss.toast('メインシートの変更を検出しました。同期処理を開始します...', '同期中', 5);
      syncMainToAllInputSheets(); // DataSync.gsの関数を呼び出し
      ss.toast('同期処理が完了しました。', '完了', 3);
    }
    // 工数シートが編集された場合
    else if (sheetName.startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
      ss.toast(`${sheetName}の変更を検出しました。集計処理を開始します...`, '集計中', 5);
      syncInputToMain(sheetName, e.range); // DataSync.gsの関数を呼び出し
      ss.toast('集計処理が完了しました。', '完了', 3);
    }
  } catch (error) {
    Logger.log(error.stack);
    ss.toast(`エラーが発生しました: ${error.message}`, "エラー", 10);
  }
}

// =================================================================================
// === カスタムメニュー機能 ===
// =================================================================================

function applyMyTasksFilter() {
  const userEmail = Session.getActiveUser().getEmail();
  const tantoushaName = getTantoushaNameByEmail(userEmail);

  if (tantoushaName) {
    const filterSettings = { tantousha: [tantoushaName], progress: [] };
    applySidebarFilters(filterSettings, true);
    SpreadsheetApp.getActiveSpreadsheet().toast(`担当者: ${tantoushaName} で絞り込みました。`);
  } else {
    SpreadsheetApp.getUi().alert('あなたのメールアドレスが担当者マスタに見つかりません。');
  }
}

function clearAllFilters() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.MAIN);
  if (!sheet) return;

  const userEmail = Session.getActiveUser().getEmail();
  const filterView = findUserFilterView(sheet, userEmail);
  if (filterView) {
    filterView.remove();
  }
  clearLastFilter();
  SpreadsheetApp.getActiveSpreadsheet().toast('個人用のフィルタを解除しました。');
}

function showFilterSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar.html')
      .setTitle('パーソナルフィルタ')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// =================================================================================
// === フィルタの内部処理 (【重要】個人用フィルタ表示) ===
// =================================================================================

/**
 * ユーザーごとのフィルタ表示(Filter View)を検索します。
 */
function findUserFilterView(sheet, userEmail) {
  const filterViews = sheet.getFilterViews();
  return filterViews.find(view => view.getName() === `FilterView-${userEmail}`);
}

/**
 * サイドバーからの情報に基づき、個人用のフィルタ表示を適用します。
 */
function applySidebarFilters(filters, save = true) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.MAIN);
  if (!sheet) return;

  const userEmail = Session.getActiveUser().getEmail();
  const indices = getColumnIndices(sheet, MAIN_SHEET_HEADERS);
  
  // 既存の個人用フィルタ表示を探す（なければ作成）
  let filterView = findUserFilterView(sheet, userEmail);
  if (filterView) {
    // 既存のフィルタ条件をクリア
    Object.values(indices).forEach(index => filterView.removeColumnFilterCriteria(index));
  } else {
    filterView = sheet.createFilterView();
    filterView.setName(`FilterView-${userEmail}`);
  }
  
  // 新しいフィルタ条件を設定
  // 担当者
  if (filters.tantousha && filters.tantousha.length > 0) {
    const criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(filters.tantousha).build();
    filterView.setColumnFilterCriteria(indices.TANTOUSHA, criteria);
  }
  // 進捗
  if (filters.progress && filters.progress.length > 0) {
    const criteria = SpreadsheetApp.newFilterCriteria().setVisibleValues(filters.progress).build();
    filterView.setColumnFilterCriteria(indices.PROGRESS, criteria);
  }

  if (save) {
    saveLastFilter(filters);
  }
}


function saveLastFilter(filters) {
  PropertiesService.getUserProperties().setProperty('lastFilter', JSON.stringify(filters));
}

function restoreLastFilter() {
  const lastFilterJson = PropertiesService.getUserProperties().getProperty('lastFilter');
  if (lastFilterJson) {
    const lastFilter = JSON.parse(lastFilterJson);
    applySidebarFilters(lastFilter, false); // 復元時は再保存しない
    SpreadsheetApp.getActiveSpreadsheet().toast('前回のフィルタを復元しました。');
  }
}

function clearLastFilter() {
  PropertiesService.getUserProperties().deleteProperty('lastFilter');
}

// =================================================================================
// === HTMLサービス連携 (サイドバー用) ===
// =================================================================================

function getFilterOptions() {
  return {
    tantousha: getMasterData(CONFIG.SHEETS.TANTOUSHA_MASTER),
    progress: getMasterData(CONFIG.SHEETS.SHINCHOKU_MASTER)
  };
}


// =================================================================================
// === データ入力規則の自動設定 ===
// =================================================================================
function setupAllDataValidations() {
  const mainSheet = new MainSheet().getSheet();
  const lastRow = mainSheet.getMaxRows();
  const headerIndices = getColumnIndices(mainSheet, MAIN_SHEET_HEADERS);
  
  const validationMap = {
    [CONFIG.SHEETS.SAGYOU_KUBUN_MASTER]: headerIndices.SAGYOU_KUBUN,
    [CONFIG.SHEETS.SHINCHOKU_MASTER]: headerIndices.PROGRESS,
    [CONFIG.SHEETS.TOIAWASE_MASTER]: headerIndices.TOIAWASE,
    [CONFIG.SHEETS.TANTOUSHA_MASTER]: headerIndices.TANTOUSHA,
  };

  for (const [masterName, colIndex] of Object.entries(validationMap)) {
    if(colIndex) {
      const masterValues = getMasterData(masterName).flat(); // 1次元配列に変換
      if (masterValues.length > 0) {
        const rule = SpreadsheetApp.newDataValidation().requireValueInList(masterValues).setAllowInvalid(false).build();
        mainSheet.getRange(CONFIG.DATA_START_ROW.MAIN, colIndex, lastRow).setDataValidation(rule);
      }
    }
  }
}

/**
 * 【デバッグ用】現在操作しているユーザーのメールアドレスを確認します。
 */
function checkMyEmail() {
  const userEmail = Session.getActiveUser().getEmail();
  SpreadsheetApp.getUi().alert(`スクリプトが認識しているメールアドレスは以下です：\n\n${userEmail}`);
}

/**
 * 【デバッグ用】スクリプトが記憶しているマスタシートのキャッシュをクリアします。
 */
function clearCache() {
  const cache = CacheService.getScriptCache();
  // 問題の原因となっている可能性のあるマスタのキャッシュキーを具体的に指定して削除
  const keysToRemove = [
    `master_${CONFIG.SHEETS.TANTOUSHA_MASTER}_2`,
    `master_${CONFIG.SHEETS.TANTOUSHA_MASTER}_1`
  ];
  cache.removeAll(keysToRemove);
  SpreadsheetApp.getUi().alert('マスタデータのキャッシュをクリアしました。');
}