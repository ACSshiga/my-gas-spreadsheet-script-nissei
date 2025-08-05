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
 * カスタムメニューの作成と、前回のフィルタ状態の復元を行います。
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('カスタムメニュー')
    .addItem('自分の担当案件のみ表示', 'applyMyTasksFilter')
    .addItem('すべての案件を表示', 'clearAllFilters')
    .addSeparator()
    .addItem('サイドバーを開く (フィルタ)', 'showFilterSidebar')
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
      syncMainToAllInputSheets();
      ss.toast('同期処理が完了しました。', '完了', 3);
    }
    // 工数シートが編集された場合
    else if (sheetName.startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
      ss.toast(`${sheetName}の変更を検出しました。集計処理を開始します...`, '集計中', 5);
      syncInputToMain(sheetName, e.range);
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

  const filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }
  clearLastFilter();
  SpreadsheetApp.getActiveSpreadsheet().toast('すべてのフィルタを解除しました。');
}

function showFilterSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar.html')
      .setTitle('パーソナルフィルタ')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// =================================================================================
// === フィルタの内部処理 (個人用フィルタ表示) ===
// =================================================================================

/**
 * サイドバーからの情報に基づき、個人用のフィルタ表示を適用します。
 * @param {Object} filters - { tantousha: string[], progress: string[] }
 * @param {boolean} [save=true] - フィルタ設定を記憶するかどうか
 */
function applySidebarFilters(filters, save = true) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.MAIN);
  if (!sheet) return;

  const indices = getColumnIndices(sheet, MAIN_SHEET_HEADERS);
  const criteriaMap = new Map();

  // 担当者フィルタの条件を作成
  if (filters.tantousha && filters.tantousha.length > 0) {
    const tantoushaCriteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(['']) // 空白は非表示
      .setVisibleValues(filters.tantousha)
      .build();
    criteriaMap.set(indices.TANTOUSHA, tantoushaCriteria);
  }

  // 進捗フィルタの条件を作成
  if (filters.progress && filters.progress.length > 0) {
    const progressCriteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(['']) // 空白は非表示
      .setVisibleValues(filters.progress)
      .build();
    criteriaMap.set(indices.PROGRESS, progressCriteria);
  }

  // 既存のフィルタを一旦削除
  let filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }

  // 新しいフィルタを作成して適用
  filter = sheet.getDataRange().createFilter();
  criteriaMap.forEach((criteria, colIndex) => {
    filter.setColumnFilterCriteria(colIndex, criteria);
  });

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

/**
 * サイドバーに表示するためのマスタ情報を取得して返します。
 */
function getFilterOptions() {
  return {
    tantousha: getMasterData(CONFIG.SHEETS.TANTOUSHA_MASTER, 1),
    progress: getMasterData(CONFIG.SHEETS.SHINCHOKU_MASTER, 1)
  };
}


// =================================================================================
// === データ入力規則の自動設定 ===
// =================================================================================
/**
 * 全マスタシートを読み込み、メインシートにドロップダウンリストを設定します。
 */
function setupAllDataValidations() {
  const mainSheet = new MainSheet().getSheet();
  const lastRow = mainSheet.getLastRow();
  const headerIndices = getColumnIndices(mainSheet, MAIN_SHEET_HEADERS);
  
  // 各マスタと対応する列のマップ
  const validationMap = {
    [CONFIG.SHEETS.SAGYOU_KUBUN_MASTER]: headerIndices.SAGYOU_KUBUN,
    [CONFIG.SHEETS.SHINCHOKU_MASTER]: headerIndices.PROGRESS,
    [CONFIG.SHEETS.TOIAWASE_MASTER]: headerIndices.TOIAWASE,
    [CONFIG.SHEETS.TANTOUSHA_MASTER]: headerIndices.TANTOUSHA,
  };

  for (const [masterName, colIndex] of Object.entries(validationMap)) {
    if(colIndex) {
      const masterValues = getMasterData(masterName, 1);
      if (masterValues.length > 0) {
        const rule = SpreadsheetApp.newDataValidation().requireValueInList(masterValues).build();
        mainSheet.getRange(CONFIG.DATA_START_ROW.MAIN, colIndex, lastRow).setDataValidation(rule);
      }
    }
  }
}/**
 * Code.gs
 * イベントハンドラとカスタムメニューを管理する司令塔。
 * スプレッドシートの操作をトリガーに、各機能を呼び出します。
 */

// =================================================================================
// === イベントハンドラ (スプレッドシート操作時に自動実行) ===
// =================================================================================

/**
 * スプレッドシートが開かれたときに実行される関数です。
 * カスタムメニューの作成と、前回のフィルタ状態の復元を行います。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // カスタムメニューを作成
  ui.createMenu('カスタムメニュー')
    .addItem('自分の担当案件のみ表示', 'applyMyTasksFilter')
    .addItem('すべての案件を表示', 'clearAllFilters')
    .addSeparator()
    .addItem('サイドバーを開く (フィルタ)', 'showFilterSidebar')
    .addSeparator()
    .addItem('全機番の資料フォルダ作成', 'bulkCreateKibanFolders')
    .addItem('全機種シリーズのフォルダ作成', 'bulkCreateSeriesFolders')
    .addSeparator()
    .addItem('週次バックアップを作成', 'createWeeklyBackup')
    .addToUi();

  // 前回のフィルタ状態を復元
  restoreLastFilter();
}

/**
 * スプレッドシートが編集されたときに実行される関数です。
 * データの同期や自動更新処理を呼び出します。
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - イベントオブジェクト
 */
function onEdit(e) {
  if (!e || !e.source || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // メインシートが編集された場合：全担当者の工数シートに即時同期
  if (sheetName === CONFIG.SHEETS.MAIN) {
    // ToDo: メインシートの変更を全工数シートに同期する処理を実装
    // syncMainToAllInputSheets();
    SpreadsheetApp.getActiveSpreadsheet().toast('メインシートが更新されました。工数シートに同期します。');
  }
  // 工数シートが編集された場合：メインシートに実績や進捗を反映
  else if (sheetName.startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
    // ToDo: 工数シートの変更をメインシートに集計・反映する処理を実装
    // syncInputToMain(sheetName);
    SpreadsheetApp.getActiveSpreadsheet().toast(`${sheetName} が更新されました。メインシートに反映します。`);
  }
}

// =================================================================================
// === カスタムメニュー機能 (フィルタリング) ===
// =================================================================================

/**
 * 現在のユーザーが担当する案件のみを表示するフィルタを適用します。
 */
function applyMyTasksFilter() {
  const userEmail = Session.getActiveUser().getEmail();
  const tantoushaName = getTantoushaNameByEmail(userEmail); // 担当者マスタから名前を取得

  if (tantoushaName) {
    const filterSettings = {
      column: MAIN_SHEET_HEADERS.TANTOUSHA,
      value: tantoushaName
    };
    applyFilter(filterSettings);
    saveLastFilter(filterSettings); // フィルタ状態を記憶
    SpreadsheetApp.getActiveSpreadsheet().toast(`担当者: ${tantoushaName} で絞り込みました。`);
  } else {
    SpreadsheetApp.getUi().alert('あなたのメールアドレスが担当者マスタに見つかりません。');
  }
}

/**
 * 現在適用されているすべてのフィルタを解除し、全案件を表示します。
 */
function clearAllFilters() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.MAIN);
  if (!sheet) return;

  const existingFilter = sheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }
  
  // フィルタ表示をすべて削除
  const filterViews = sheet.getFilterViews();
  filterViews.forEach(view => view.remove());

  clearLastFilter(); // 記憶したフィルタを消去
  SpreadsheetApp.getActiveSpreadsheet().toast('すべてのフィルタを解除しました。');
}

/**
 * フィルタリング用のサイドバーを表示します。
 */
function showFilterSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar.html')
      .setTitle('パーソナルフィルタ')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}


// =================================================================================
// === フィルタの内部処理 ===
// =================================================================================

/**
 * 指定された設定でフィルタ表示を適用します。
 * @param {Object} filterSettings - { column: 'ヘッダー名', value: '値' }
 */
function applyFilter(filterSettings) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.MAIN);
  if (!sheet) return;

  const userEmail = Session.getActiveUser().getEmail();
  const indices = getColumnIndices(sheet, MAIN_SHEET_HEADERS);
  const targetColIndex = indices[Object.keys(MAIN_SHEET_HEADERS).find(key => MAIN_SHEET_HEADERS[key] === filterSettings.column)];

  if (!targetColIndex) {
    SpreadsheetApp.getUi().alert(`列「${filterSettings.column}」が見つかりません。`);
    return;
  }

  // 既存の個人用フィルタ表示を削除
  const filterViews = sheet.getFilterViews();
  filterViews.forEach(view => {
    if (view.getName().includes(userEmail)) {
      view.remove();
    }
  });

  // 新しいフィルタ表示を作成
  const filterView = sheet.createFilterView();
  filterView.setName(`Filter-${userEmail}-${new Date().getTime()}`); // ユニークな名前
  
  const criteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(['']) // 空白行は非表示
      .whenTextEqualTo(filterSettings.value)
      .build();
  
  filterView.setColumnFilterCriteria(targetColIndex, criteria);
  
  // 重要：フィルタ表示を適用する (これだけでは他の人に見えない)
  SpreadsheetApp.setActiveSheet(sheet); // シートをアクティブに
  SpreadsheetApp.flush(); // 保留中の変更を適用
  // 注意: Apps Scriptから直接FilterViewを「アクティブ」にするAPIはないため、
  // ユーザーに手動で適用してもらうか、このスクリプトでは設定の保存と作成に注力します。
  // ここでは、フィルタを作成し、ユーザーが手動でオンにできるように準備します。
}


/**
 * 最後に使用したフィルタ設定をユーザープロパティに保存します。
 * @param {Object} filterSettings - 保存するフィルタ設定
 */
function saveLastFilter(filterSettings) {
  PropertiesService.getUserProperties().setProperty('lastFilter', JSON.stringify(filterSettings));
}

/**
 * 保存されたフィルタ設定を読み込み、シートに適用します。
 */
function restoreLastFilter() {
  const lastFilterJson = PropertiesService.getUserProperties().getProperty('lastFilter');
  if (lastFilterJson) {
    const lastFilter = JSON.parse(lastFilterJson);
    applyFilter(lastFilter);
    SpreadsheetApp.getActiveSpreadsheet().toast('前回のフィルタを復元しました。');
  }
}

/**
 * 保存されたフィルタ設定を削除します。
 */
function clearLastFilter() {
  PropertiesService.getUserProperties().deleteProperty('lastFilter');
}

/**
* メールアドレスから担当者マスタを検索し、対応する担当者名を返します。
* @param {string} email - 検索するメールアドレス
* @return {string|null} - 見つかった担当者名、またはnull
*/
function getTantoushaNameByEmail(email) {
  const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.TANTOUSHA_MASTER);
  if (!masterSheet) return null;

  const data = masterSheet.getDataRange().getValues();
  // ヘッダーをスキップして検索
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === email) { // 2列目(B列)にメールアドレスがあると仮定
      return data[i][0]; // 1列目(A列)の担当者名を返す
    }
  }
  return null;
}/**
 * Code.gs
 * イベントハンドラとカスタムメニューを管理する司令塔。
 */

function onEdit(e) {
  if (!e || !e.source || !e.range) { return; }
  
  const ss = e.source;
  const sheetName = e.range.getSheet().getName();
  
  ss.toast(`処理を開始します... (${sheetName}シート編集中)`, "情報", 3);

  try {
    // 編集があった場合は、常に全シートの更新を実行するシンプルな設計
    updateAllSheets();
    ss.toast("自動処理が完了しました。", "完了", 3);

  } catch (error) {
    Logger.log(error.stack);
    ss.toast(`エラーが発生しました: ${error.message}`, "エラー", 10);
  }
}

/**
 * すべてのシートのデータ整合性と書式を更新する統合関数。
 */
function updateAllSheets() {
  logWithTimestamp("全シートの更新処理を開始します。");
  
  const mainSheet = new MainSheet();
  const tantoushaList = mainSheet.getTantoushaList();
  
  // 1. 全工数シートから実績工数を集計し、メインシートに反映
  mainSheet.updateActualHours();
  
  // 2. メインシートの最新データを全工数シートに同期（再構築）
  const mainDataMap = mainSheet.getDataMap();
  tantoushaList.forEach(tantousha => {
    try {
      const inputSheet = new InputSheet(tantousha);
      inputSheet.syncFromMain(mainDataMap);
    } catch (e) {}
  });
  
  // 3. 全シートの色付けを実行
  colorizeAllSheets();
  
  logWithTimestamp("全シートの更新処理が完了しました。");
}

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
}/**
 * Code.gs
 * イベントハンドラとカスタムメニューを管理する司令塔。
 * オブジェクト指向設計に基づき、各シートクラスのインスタンスを呼び出す役割を担う。
 */

// =_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
// === イベントハンドラ (司令塔) ===
// =_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

/**
 * スプレッドシートが編集されたときに自動的に実行されるトリガー関数。
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - イベントオブジェクト
 */
function onEdit(e) {
  if (!e || !e.source || !e.range) { return; }
  
  const ss = e.source;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  
  ss.toast(`処理を開始します... (${sheetName}シート編集中)`, "情報", 3);

  try {
    const mainSheet = new MainSheet();
    const allTantousha = mainSheet.getTantoushaList();

    // ===============================================================
    // === メインシートが編集された場合の処理 ===
    // ===============================================================
    if (sheetName === CONFIG.SHEETS.MAIN) {
      // 担当者マスタに存在する担当者の工数シートのみを同期対象とする
      allTantousha.forEach(tantousha => {
        try {
          const inputSheet = new InputSheet(tantousha);
          // ここに同期処理を実装
        } catch (error) {
          // シートが存在しない場合は何もしない
        }
      });
    
    // ===============================================================
    // === 工数シートが編集された場合の処理 ===
    // ===============================================================
    } else if (sheetName.startsWith(CONFIG.SHEETS.INPUT_PREFIX)) {
      const tantoushaName = sheetName.replace(CONFIG.SHEETS.INPUT_PREFIX, '');
      if (allTantousha.includes(tantoushaName)) {
        const inputSheet = new InputSheet(tantoushaName);
        // ここにメインシートへの同期処理を実装
      }
    }

    // すべての処理の最後に全体を更新・整形
    updateAllSheets();

    ss.toast("自動処理が完了しました。", "完了", 3);

  } catch (error) {
    Logger.log(error.stack);
    ss.toast(`エラーが発生しました: ${error.message}`, "エラー", 10);
  }
}

/**
 * すべてのシートのデータ整合性と書式を更新する統合関数。
 */
function updateAllSheets() {
  // ここで再構築、実績工数の集計、重複チェック、色付けなどを一括して行う
  // 詳細な実装は今後のステップで追加します。
  logWithTimestamp("全シートの更新処理を開始します。");
  
  // 1. 実績工数をメインシートに集計
  // 2. メインシートの情報を各工数シートに同期
  // 3. 重複チェック
  // 4. 色付け
  
  logWithTimestamp("全シートの更新処理が完了しました。");
}


// =_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
// === カスタムメニュー ===
// =_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

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
}/**
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