/**
 * ★★★ 修正箇所あり ★★★
 * 指定されたシートオブジェクトの各列をルールに基づき色付けします。
 * 数式を保護しながら、重複行（同じ機番＋作業区分）の検出と処理を行います。
 */
function colorizeSheet_(sheetObject) {
  const sheet = sheetObject.getSheet();
  const indices = sheetObject.indices;
  const lastRow = sheetObject.getLastRow();
  const startRow = sheetObject.startRow;

  if (lastRow < startRow) return;
  const dataRows = lastRow - startRow + 1;
  const lastCol = sheet.getLastColumn();
  const fullRange = sheet.getRange(startRow, 1, dataRows, lastCol);
  const displayValues = fullRange.getDisplayValues();
  const formulas = fullRange.getFormulas();
  const backgroundColors = fullRange.getBackgrounds();
  
  const outputValues = JSON.parse(JSON.stringify(displayValues));

  const mgmtNoCol = indices.MGMT_NO;
  const progressCol = indices.PROGRESS;
  const tantoushaCol = indices.TANTOUSHA;
  const toiawaseCol = indices.TOIAWASE;
  const kibanCol = indices.KIBAN;
  const sagyouKubunCol = indices.SAGYOU_KUBUN; // 作業区分列のインデックスを取得
  const DUPLICATE_COLOR = '#cccccc';

  const uniqueKeys = new Set();
  const restrictedRanges = [];
  const normalRanges = [];
  displayValues.forEach((row, i) => {
    let isDuplicate = false;
    if (kibanCol && sagyouKubunCol) {
      const kiban = safeTrim(row[kibanCol - 1]);
      const sagyouKubun = safeTrim(row[sagyouKubunCol - 1]);
      
      if (kiban && sagyouKubun) {
        const uniqueKey = `${kiban}_${sagyouKubun}`;
        if (uniqueKeys.has(uniqueKey)) {
          isDuplicate = true;
        } else {
          uniqueKeys.add(uniqueKey);
        }
      }
    }
    
    if (isDuplicate) {
      for (let j = 0; j < lastCol; j++) {
        backgroundColors[i][j] = DUPLICATE_COLOR;
      }
      if (progressCol) outputValues[i][progressCol - 1] = "機番重複";
      if (tantoushaCol) outputValues[i][tantoushaCol - 1] = "";
      if (sheetObject instanceof MainSheet && tantoushaCol) {
        restrictedRanges.push(sheet.getRange(startRow + i, tantoushaCol));
      }
    } else {
      // 既存の色をリセット
      for (let j = 0; j < lastCol; j++) {
          if(!formulas[i][j]) backgroundColors[i][j] = CONFIG.COLORS.DEFAULT_BACKGROUND;
      }

      if (progressCol) {
        const progressColor = getColor(PROGRESS_COLORS, safeTrim(row[progressCol - 1]));
        backgroundColors[i][progressCol - 1] = progressColor;
        if (mgmtNoCol) backgroundColors[i][mgmtNoCol - 1] = progressColor;
      }
      
      // ★★★ ここから追加 ★★★
      if (sagyouKubunCol) {
        const sagyouKubunColor = getColor(SAGYOU_KUBUN_COLORS, safeTrim(row[sagyouKubunCol - 1]));
        backgroundColors[i][sagyouKubunCol - 1] = sagyouKubunColor;
      }
      // ★★★ ここまで追加 ★★★

      if (sheetObject instanceof MainSheet || sheet.getName().startsWith('View_')) {
        if (tantoushaCol) backgroundColors[i][tantoushaCol - 1] = getColor(TANTOUSHA_COLORS, safeTrim(row[tantoushaCol - 1]));
        if (toiawaseCol) backgroundColors[i][toiawaseCol - 1] = getColor(TOIAWASE_COLORS, safeTrim(row[toiawaseCol - 1]));
      }
      if (sheetObject instanceof MainSheet && tantoushaCol) {
        normalRanges.push(sheet.getRange(startRow + i, tantoushaCol));
      }
    }
  });

  // (以下、変更なし)
  formulas.forEach((row, i) => {
    row.forEach((formula, j) => {
      if (formula) {
        outputValues[i][j] = formula;
      }
    });
  });
  fullRange.setBackgrounds(backgroundColors);
  fullRange.setValues(outputValues);

  if (sheetObject instanceof MainSheet) {
    if (restrictedRanges.length > 0) {
      const restrictedRule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(sheet.getRange('A1:A1'), false)
        .setAllowInvalid(false)
        .setHelpText('この行は機番が重複しているため、担当者は設定できません。')
        .build();
      restrictedRanges.forEach(range => range.setDataValidation(restrictedRule));
    }
    if (normalRanges.length > 0) {
      const masterValues = getMasterData(CONFIG.SHEETS.TANTOUSHA_MASTER).flat();
      if (masterValues.length > 0) {
        const normalRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(masterValues)
          .setAllowInvalid(false)
          .build();
        normalRanges.forEach(range => range.setDataValidation(normalRule));
      } else {
         normalRanges.forEach(range => range.clearDataValidations());
      }
    }
  }
}