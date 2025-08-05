/**
 * SheetService.gs
 * スプレッドシートの各シートをオブジェクトとして効率的に扱うためのクラスを定義します。
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
      this.sheet = this.ss.insertSheet(sheetName);
    }
    this.sheetName = sheetName;
    this.startRow = 2;
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
    this.startRow = CONFIG.DATA_START_ROW.MAIN;
    this.indices = getColumnIndices(this.sheet, MAIN_SHEET_HEADERS);
  }

  getTantoushaList() {
    return getMasterData(CONFIG.SHEETS.TANTOUSHA_MASTER, 2)
      .map(row => ({ name: row[0], email: row[1] }))
      .filter(item => item.name && item.email);
  }

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
  constructor(tantoushaName) {
    const sheetName = `${CONFIG.SHEETS.INPUT_PREFIX}${tantoushaName}`;
    super(sheetName);
    this.tantoushaName = tantoushaName;
    this.startRow = CONFIG.DATA_START_ROW.INPUT;
    
    if (this.sheet.getLastRow() < 2) {
      this.initializeSheet();
    }
    this.indices = getColumnIndices(this.sheet, INPUT_SHEET_HEADERS);
    
    this.filterDateColumns();
  }

  initializeSheet() {
    this.sheet.clear();
    const headers = Object.values(INPUT_SHEET_HEADERS);
    this.sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    
    const separatorCol = headers.indexOf("");
    if (separatorCol !== -1) {
      this.sheet.getRange(2, separatorCol + 1).setValue("日次合計").setHorizontalAlignment("right");
    }

    const dateHeaders = [];
    const sumFormulas = [];
    const today = new Date();
    const prevMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    const thisMonth = new Date(today.getFullYear(), today.getMonth(), 1);
    const datesToGenerate = [prevMonth, thisMonth];
    let currentCol = headers.length + 1;
    datesToGenerate.forEach(date => {
      const year = date.getFullYear();
      const month = date.getMonth();
      const daysInMonth = new Date(year, month + 1, 0).getDate();
      for (let i = 1; i <= daysInMonth; i++) {
        dateHeaders.push(new Date(year, month, i));
        const colLetter = this.sheet.getRange(1, currentCol).getA1Notation().replace("1", "");
        sumFormulas.push(`=IFERROR(SUM(${colLetter}${this.startRow}:${colLetter}))`);
        currentCol++;
      }
    });
    if (dateHeaders.length > 0) {
      this.sheet.getRange(1, headers.length + 1, 1, dateHeaders.length).setValues([dateHeaders]).setNumberFormat("M/d");
      this.sheet.getRange(2, headers.length + 1, 1, sumFormulas.length).setFormulas([sumFormulas]);
    }

    this.sheet.setFrozenRows(2);
    this.sheet.setFrozenColumns(7);
  }
  
  filterDateColumns() {
    const sheet = this.sheet;
    const lastCol = sheet.getLastColumn();
    const dateStartCol = Object.keys(INPUT_SHEET_HEADERS).length + 1;

    if (lastCol < dateStartCol) return;

    sheet.showColumns(dateStartCol, lastCol - dateStartCol + 1);
    const today = new Date();
    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth();
    const headerDates = sheet.getRange(1, dateStartCol, 1, lastCol - dateStartCol + 1).getValues()[0];
    headerDates.forEach((date, i) => {
      if (isValidDate(date)) {
        const dateYear = date.getFullYear();
        const dateMonth = date.getMonth();
        
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

  // ▼▼▼ ここから修正 ▼▼▼
  clearData() {
    const existingFilter = this.sheet.getFilter();
    if (existingFilter) {
      existingFilter.remove();
    }
    const lastRow = this.getLastRow();
    if (lastRow >= this.startRow) {
      // .clearContent() から .clear() に変更して、書式ごとクリアする
      this.sheet.getRange(this.startRow, 1, lastRow - this.startRow + 1, this.getLastColumn()).clear();
    }
  }
  // ▲▲▲ ここまで修正 ▲▲▲

  writeData(data) {
    const existingFilter = this.sheet.getFilter();
    if (existingFilter) {
      existingFilter.remove();
    }
    
    if (data.length === 0) return;
    this.sheet.getRange(this.startRow, 1, data.length, data[0].length).setValues(data);

    const sumFormulas = [];
    const dateStartCol = Object.keys(INPUT_SHEET_HEADERS).length + 1;
    const dateStartColLetter = this.sheet.getRange(1, dateStartCol).getA1Notation().replace("1", "");
    for (let i = 0; i < data.length; i++) {
      const rowNum = this.startRow + i;
      sumFormulas.push([`=IFERROR(SUM(${dateStartColLetter}${rowNum}:${rowNum}))`]);
    }
    this.sheet.getRange(this.startRow, this.indices.ACTUAL_HOURS_SUM, data.length, 1).setFormulas(sumFormulas);
    
    if(this.sheet.getLastRow() > 1) {
      this.sheet.getRange(1, 1, this.sheet.getLastRow(), this.sheet.getLastColumn()).createFilter();
    }
  }
}