/**
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
}