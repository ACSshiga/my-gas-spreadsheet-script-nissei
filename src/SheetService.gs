/**
 * SheetService.gs
 * スプレッドシートの各シートをオブジェクトとして効率的に扱うためのクラスを定義します。
 */

// =================================================================================
// === すべてのシートクラスの基盤となる抽象クラス ===
// =================================================================================
class SheetService {
  // (変更なしのため省略)
}


// =================================================================================
// === メインシートを操作するためのクラス ===
// =================================================================================
class MainSheet extends SheetService {
  // (変更なしのため省略)
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
    
    const sumLabelCol = headers.indexOf("実績工数合計");
    if (sumLabelCol > 0) {
      this.sheet.getRange(2, sumLabelCol).setValue("日次合計").setHorizontalAlignment("right");
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
    this.sheet.setFrozenColumns(6); // ★固定列をF列（6列）までに変更
  }
  
  filterDateColumns() {
    // (変更なしのため省略)
  }

  clearData() {
    // (変更なしのため省略)
  }

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
    
    // ★★★ 修正箇所 ★★★
    // ヘッダー行(1行目)からデータ最終行までを範囲としてフィルタを作成
    if(this.sheet.getLastRow() > 1) {
      this.sheet.getRange(1, 1, this.sheet.getLastRow(), this.sheet.getLastColumn()).createFilter();
    }
  }
}