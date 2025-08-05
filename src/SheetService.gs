/**
 * SheetService.gs
 * スプレッドシートの各シートをオブジェクトとして効率的に扱うためのクラスを定義します。
 * この設計により、コードの再利用性とメンテナンス性が向上します。
 */

// =================================================================================
// === すべてのシートクラスの基盤となる抽象クラス ===
// =================================================================================
class SheetService {
  /**
   * @param {string} sheetName 操作対象のシート名
   */
  constructor(sheetName) {
    if (this.constructor === SheetService) {
      throw new Error("SheetServiceは抽象クラスのためインスタンス化できません。");
    }
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.sheet = this.ss.getSheetByName(sheetName);
    if (!this.sheet) {
      // シートが存在しない場合は新規作成する
      this.sheet = this.ss.insertSheet(sheetName);
    }
    this.sheetName = sheetName;
    this.startRow = 2; // デフォルトのデータ開始行
  }

  /** @return {GoogleAppsScript.Spreadsheet.Sheet} シートオブジェクトを返す */
  getSheet() { return this.sheet; }
  /** @return {number} シートの最終行番号を返す */
  getLastRow() { return this.sheet.getLastRow(); }
  /** @return {number} シートの最終列番号を返す */
  getLastColumn() { return this.sheet.getLastColumn(); }
  /** @return {string} シート名を返す */
  getName() { return this.sheetName; }
}


// =================================================================================
// === メインシートを操作するためのクラス ===
// =================================================================================
class MainSheet extends SheetService {
  constructor() {
    super(CONFIG.SHEETS.MAIN);
    this.startRow = CONFIG.DATA_START_ROW.MAIN;
    this.indices = getColumnIndices(this.sheet, MAIN_SHEET_HEADERS);
  }

  /**
   * 「担当者マスタ」シートから担当者の情報（名前とメールアドレス）を取得します。
   */
  getTantoushaList() {
    return getMasterData(CONFIG.SHEETS.TANTOUSHA_MASTER, 2)
      .map(row => ({ name: row[0], email: row[1] }))
      .filter(item => item.name && item.email);
  }

  /**
   * メインシートの全データを、一意なキー（管理No + 作業区分）でMapとして取得します。
   */
  getDataMap() {
    const lastRow = this.getLastRow();
    if (lastRow < this.startRow) return new Map();

    const values = this.sheet.getRange(this.startRow, 1, lastRow - this.startRow + 1, this.getLastColumn()).getValues();
    const dataMap = new Map();

    values.forEach((row, index) => {
      const mgmtNo = row[this.indices.MGMT_NO - 1];
      const sagyouKubun = row[this.indices.SAGYOU_KUBUN - 1];
      if (mgmtNo && sagyouKubun) {
        const uniqueKey = `${mgmtNo}_${sagyouKubun}`;
        dataMap.set(uniqueKey, { data: row, rowNum: this.startRow + index });
      }
    });
    return dataMap;
  }
}


// =================================================================================
// === 工数シートを操作するためのクラス ===
// =================================================================================
class InputSheet extends SheetService {
  /**
   * @param {string} tantoushaName 担当者名
   */
  constructor(tantoushaName) {
    const sheetName = `${CONFIG.SHEETS.INPUT_PREFIX}${tantoushaName}`;
    super(sheetName);
    this.tantoushaName = tantoushaName;
    this.startRow = CONFIG.DATA_START_ROW.INPUT;
    this.indices = getColumnIndices(this.sheet, INPUT_SHEET_HEADERS);

    // シートが新規作成された場合、または空の場合に初期化
    if (this.sheet.getLastRow() < 2) {
      this.initializeSheet();
    }
    
    // ★表示する日付列を当月と前月に絞る
    this.filterDateColumns();
  }

  /**
   * 工数シートを初期化し、ヘッダーと数式を設定します。
   */
  initializeSheet() {
    this.sheet.clear(); // シートをクリア
    const headers = Object.values(INPUT_SHEET_HEADERS);
    this.sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    this.sheet.getRange(2, headers.length - 1).setValue("日次合計").setHorizontalAlignment("right");
    
    this.sheet.setFrozenRows(2);
    this.sheet.setFrozenColumns(3);
    this.updateDateColumns(); // 日付列を生成
  }
  
  /**
   * ★工数シートの日付列を更新・生成します。
   */
  updateDateColumns() {
      // ToDo: 必要に応じて日付列を動的に追加・更新するロジックを実装
  }
  
  /**
   * ★表示する日付列を、当月と前月のものだけにフィルタリングします。
   */
  filterDateColumns() {
    const sheet = this.sheet;
    const lastCol = sheet.getLastColumn();
    const dateStartCol = Object.keys(INPUT_SHEET_HEADERS).length + 1;

    if (lastCol < dateStartCol) return;

    sheet.showColumns(dateStartCol, lastCol - dateStartCol + 1); // 一旦すべて表示

    const today = new Date();
    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth(); // 0-indexed

    const headerDates = sheet.getRange(1, dateStartCol, 1, lastCol - dateStartCol + 1).getValues()[0];
    
    headerDates.forEach((date, i) => {
      if (isValidDate(date)) {
        const dateYear = date.getFullYear();
        const dateMonth = date.getMonth();
        
        // 当月または前月でない場合は非表示
        const isCurrentMonth = (dateYear === currentYear && dateMonth === currentMonth);
        const isPreviousMonth = (currentMonth === 0) 
            ? (dateYear === currentYear - 1 && dateMonth === 11) 
            : (dateYear === currentYear && dateMonth === currentMonth - 1);

        if (!isCurrentMonth && !isPreviousMonth) {
          sheet.hideColumns(dateStartCol + i);
        }
      }
    });
  }


  /**
   * 工数シートの既存データをクリアします。
   */
  clearData() {
    const lastRow = this.getLastRow();
    if (lastRow >= this.startRow) {
      this.sheet.getRange(this.startRow, 1, lastRow - this.startRow + 1, this.getLastColumn()).clearContent();
    }
  }

  /**
   * 新しいデータを工数シートに書き込み、数式も設定します。
   */
  writeData(data) {
    if (data.length === 0) return;
    
    this.sheet.getRange(this.startRow, 1, data.length, data[0].length).setValues(data);

    // 実績工数合計の数式を設定
    const sumFormulas = [];
    const dateStartCol = Object.keys(INPUT_SHEET_HEADERS).length + 1;
    const dateStartColLetter = this.sheet.getRange(1, dateStartCol).getA1Notation().replace("1", "");
    
    for (let i = 0; i < data.length; i++) {
      const rowNum = this.startRow + i;
      sumFormulas.push([`=SUM(${dateStartColLetter}${rowNum}:${rowNum})`]);
    }
    this.sheet.getRange(this.startRow, this.indices.ACTUAL_HOURS_SUM, data.length, 1).setFormulas(sumFormulas);
  }
}