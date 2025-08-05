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
      throw new Error(`シート「${sheetName}」が見つかりません。`);
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
   * @return {{name: string, email: string}[]} 担当者情報の配列
   */
  getTantoushaList() {
    const masterSheet = this.ss.getSheetByName(CONFIG.SHEETS.TANTOUSHA_MASTER);
    if (!masterSheet) {
      SpreadsheetApp.getUi().alert(`シート「${CONFIG.SHEETS.TANTOUSHA_MASTER}」が見つかりません。`);
      return [];
    }
    const lastRow = masterSheet.getLastRow();
    if (lastRow < 2) return [];

    const data = masterSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    return data.map(row => ({ name: row[0], email: row[1] })).filter(item => item.name && item.email);
  }

  /**
   * メインシートの全データを、一意なキー（管理No + 作業区分）でMapとして取得します。
   * @return {Map<string, Object>} キーを '管理No_作業区分', 値を行データ配列とするMap
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
   * @param {string} tantoushaName 担当者名 (例: "志賀")
   */
  constructor(tantoushaName) {
    const sheetName = `${CONFIG.SHEETS.INPUT_PREFIX}${tantoushaName}`;
    super(sheetName);
    this.tantoushaName = tantoushaName;
    this.startRow = CONFIG.DATA_START_ROW.INPUT;
    this.indices = getColumnIndices(this.sheet, INPUT_SHEET_HEADERS);
  }

  /**
   * このシートの担当者名を返します。
   * @return {string}
   */
  getTantoushaName() {
    return this.tantoushaName;
  }
  
  /**
   * 工数シートの既存データをクリアします（ヘッダーと日次合計行は除く）。
   */
  clearData() {
    const lastRow = this.getLastRow();
    if (lastRow >= this.startRow) {
      this.sheet.getRange(this.startRow, 1, lastRow - this.startRow + 1, this.getLastColumn()).clearContent();
    }
  }

  /**
   * 新しいデータを工数シートに書き込みます。
   * @param {Array<Array<any>>} data - 書き込む2次元配列データ
   */
  writeData(data) {
    if (data.length === 0) return;
    this.sheet.getRange(this.startRow, 1, data.length, data[0].length).setValues(data);
  }
}/**
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
      throw new Error(`シート「${sheetName}」が見つかりません。`);
    }
    this.sheetName = sheetName;
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
    this.indices = getColumnIndices(this.sheet, MAIN_SHEET_HEADERS);
  }

  /**
   * 「担当者マスタ」シートから担当者の情報（名前とメールアドレス）を取得します。
   * @return {{name: string, email: string}[]} 担当者情報の配列
   */
  getTantoushaList() {
    const masterSheet = this.ss.getSheetByName(CONFIG.SHEETS.TANTOUSHA_MASTER);
    if (!masterSheet) {
      SpreadsheetApp.getUi().alert(`シート「${CONFIG.SHEETS.TANTOUSHA_MASTER}」が見つかりません。`);
      return [];
    }
    const lastRow = masterSheet.getLastRow();
    if (lastRow < 2) return [];

    const data = masterSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    return data.map(row => ({ name: row[0], email: row[1] })).filter(item => item.name && item.email);
  }

  /**
   * メインシートの全データを、一意なキー（管理No + 作業区分）でMapとして取得します。
   * @return {Map<string, Object>}
   */
  getDataMap() {
    const lastRow = this.getLastRow();
    if (lastRow < 2) return new Map();

    const values = this.sheet.getRange(2, 1, lastRow - 1, this.getLastColumn()).getValues();
    const dataMap = new Map();

    values.forEach(row => {
      const mgmtNo = row[this.indices.MGMT_NO - 1];
      const sagyouKubun = row[this.indices.SAGYOU_KUBUN - 1];
      if (mgmtNo && sagyouKubun) {
        const uniqueKey = `${mgmtNo}_${sagyouKubun}`;
        dataMap.set(uniqueKey, row);
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
   * @param {string} tantoushaName 担当者名 (例: "志賀")
   */
  constructor(tantoushaName) {
    const sheetName = `${CONFIG.SHEETS.INPUT_PREFIX}${tantoushaName}`;
    super(sheetName);
    this.tantoushaName = tantoushaName;
    this.indices = getColumnIndices(this.sheet, INPUT_SHEET_HEADERS);
  }

  /**
   * このシートの担当者名を返します。
   * @return {string}
   */
  getTantoushaName() {
    return this.tantoushaName;
  }
}/**
 * SheetService.gs
 * スプレッドシートの各シートをオブジェクトとして扱うためのクラスを定義します。
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
      throw new Error(`シート「${sheetName}」が見つかりません。`);
    }
    this.sheetName = sheetName;
  }

  getSheet() { return this.sheet; }
  getLastRow() { return this.sheet.getLastRow(); }
  getLastColumn() { return this.sheet.getLastColumn(); }
  getName() { return this.sheetName; }
}


// =================================================================================
// === メインシートを操作するためのクラス ===
// =================================================================================
class MainSheet extends SheetService {
  constructor() {
    super(CONFIG.SHEETS.MAIN);
    this.indices = getColumnIndices(this.sheet, MAIN_SHEET_HEADERS);
  }

  /**
   * 「担当者マスタ」シートから担当者のリストを取得します。
   * @return {string[]} 担当者名の配列
   */
  getTantoushaList() {
    const masterSheet = this.ss.getSheetByName(CONFIG.SHEETS.TANTOUSHA_MASTER);
    if (!masterSheet) {
      throw new Error(`シート「${CONFIG.SHEETS.TANTOUSHA_MASTER}」が見つかりません。`);
    }
    const lastRow = masterSheet.getLastRow();
    if (lastRow < 2) return [];
    
    const names = masterSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    return names.flat().filter(name => name);
  }

  /**
   * メインシートの全データを管理NoをキーにしたMapとして取得します。
   * @return {Map<string, Object>}
   */
  getDataMap() {
    const lastRow = this.getLastRow();
    if (lastRow < 2) return new Map();
    const values = this.sheet.getRange(2, 1, lastRow - 1, this.getLastColumn()).getValues();
    return new Map(values.map(row => [safeTrim(row[this.indices.MGMT_NO - 1]), row]));
  }
}


// =================================================================================
// === 工数シートを操作するためのクラス ===
// =================================================================================
class InputSheet extends SheetService {
  /**
   * @param {string} tantoushaName 担当者名 (例: "志賀")
   */
  constructor(tantoushaName) {
    const sheetName = `${CONFIG.SHEETS.INPUT_PREFIX}${tantoushaName}`;
    super(sheetName);
    this.tantoushaName = tantoushaName;
    this.indices = getColumnIndices(this.sheet, INPUT_SHEET_HEADERS);
  }

  /**
   * このシートの担当者名を返します。
   * @return {string}
   */
  getTantoushaName() {
    return this.tantoushaName;
  }
}/**
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
      // 新規作成を試みる
      this.sheet = this.ss.insertSheet(sheetName);
      // ここでヘッダーなどを初期設定することも可能
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
   * @return {{name: string, email: string}[]} 担当者情報の配列
   */
  getTantoushaList() {
    return getMasterData(CONFIG.SHEETS.TANTOUSHA_MASTER, 2)
      .map(row => ({ name: row[0], email: row[1] }))
      .filter(item => item.name && item.email);
  }

  /**
   * メインシートの全データを、一意なキー（管理No + 作業区分）でMapとして取得します。
   * @return {Map<string, {data: any[], rowNum: number}>} キーを '管理No_作業区分', 値を {データ配列, 行番号} とするMap
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
   * @param {string} tantoushaName 担当者名 (例: "志賀")
   */
  constructor(tantoushaName) {
    const sheetName = `${CONFIG.SHEETS.INPUT_PREFIX}${tantoushaName}`;
    super(sheetName);
    this.tantoushaName = tantoushaName;
    this.startRow = CONFIG.DATA_START_ROW.INPUT;
    this.indices = getColumnIndices(this.sheet, INPUT_SHEET_HEADERS);

    // シートが新規作成された場合、ヘッダーと数式を初期設定
    if (this.sheet.getLastRow() < 2) {
      this.initializeSheet();
    }
  }

  /**
   * 工数シートを初期化し、ヘッダーと数式を設定します。
   */
  initializeSheet() {
    const headers = Object.values(INPUT_SHEET_HEADERS);
    this.sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.sheet.getRange("F2").setValue("日次合計").setHorizontalAlignment("right");
    
    const dateHeaderRange = this.sheet.getRange(1, headers.length + 1, 1, 31);
    const dailySumFormulaRange = this.sheet.getRange(2, headers.length + 1, 1, 31);

    // 今月の日付ヘッダーと日次合計数式を設定
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    
    const dateHeaders = [];
    const sumFormulas = [];
    for (let i = 1; i <= daysInMonth; i++) {
      dateHeaders.push(new Date(year, month, i));
      const colLetter = this.sheet.getRange(1, headers.length + i).getA1Notation().replace("1", "");
      sumFormulas.push(`=SUM(${colLetter}${this.startRow}:${colLetter})`);
    }
    dateHeaderRange.setValues([dateHeaders]).setNumberFormat("M/d");
    dailySumFormulaRange.setFormulas([sumFormulas]);

    this.sheet.setFrozenRows(2);
    this.sheet.setFrozenColumns(3);
  }

  /**
   * 工数シートの既存データをクリアします（ヘッダーと日次合計行は除く）。
   */
  clearData() {
    const lastRow = this.getLastRow();
    if (lastRow >= this.startRow) {
      this.sheet.getRange(this.startRow, 1, lastRow - this.startRow + 1, this.getLastColumn()).clearContent();
    }
  }

  /**
   * 新しいデータを工数シートに書き込み、実績工数合計の数式も設定します。
   * @param {Array<Array<any>>} data - 書き込む2次元配列データ
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