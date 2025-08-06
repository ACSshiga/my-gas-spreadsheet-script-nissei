/**
 * DataSync.gs
 * メインシートと工数シート間のデータ同期処理を専門に担当します。
 */

// =================================================================================
// === メイン → 工数シートへの同期処理 (差分更新アプローチ) ===
// =================================================================================
function syncMainToAllInputSheets() {
  const mainSheet = new MainSheet();
  const mainIndices = mainSheet.indices;

  // --- メインシートの「未着手」自動入力 ---
  if (mainSheet.getLastRow() >= mainSheet.startRow) {
    const range = mainSheet.sheet.getRange(mainSheet.startRow, 1, mainSheet.getLastRow() - mainSheet.startRow + 1, mainSheet.getLastColumn());
    const mainData = range.getValues();
    let hasUpdate = false;
    mainData.forEach(row => {
      if (!row[mainIndices.PROGRESS - 1] && row[mainIndices.TANTOUSHA - 1]) {
        row[mainIndices.PROGRESS - 1] = "未着手";
        hasUpdate = true;
      }
    });
    if (hasUpdate) range.setValues(mainData);
  }
  
  const mainDataValues = mainSheet.sheet.getRange(mainSheet.startRow, 1, mainSheet.getLastRow() - mainSheet.startRow + 1, mainSheet.getLastColumn()).getValues();
  
  // --- 担当者ごとに、メインシートにあるべき案件リストを作成 ---
  const projectsByTantousha = new Map();
  mainDataValues.forEach(row => {
    const tantousha = row[mainIndices.TANTOUSHA - 1];
    if (!tantousha) return;
    if (!projectsByTantousha.has(tantousha)) {
      projectsByTantousha.set(tantousha, new Map());
    }
    const key = `${row[mainIndices.MGMT_NO - 1]}_${row[mainIndices.SAGYOU_KUBUN - 1]}`;
    projectsByTantousha.get(tantousha).set(key, row);
  });

  // --- 各担当者シートをチェックして差分更新 ---
  const tantoushaList = mainSheet.getTantoushaList();
  tantoushaList.forEach(tantousha => {
    try {
      const inputSheet = new InputSheet(tantousha.name);
      const mainProjects = projectsByTantousha.get(tantousha.name) || new Map();
      
      const inputSheetRange = (inputSheet.getLastRow() >= inputSheet.startRow) 
        ? inputSheet.sheet.getRange(inputSheet.startRow, 1, inputSheet.getLastRow() - inputSheet.startRow + 1, inputSheet.getLastColumn())
        : null;
      const inputValues = inputSheetRange ? inputSheetRange.getValues() : [];
      
      const inputProjects = new Map();
      inputValues.forEach((row, i) => {
        const key = `${row[inputSheet.indices.MGMT_NO - 1]}_${row[inputSheet.indices.SAGYOU_KUBUN - 1]}`;
        inputProjects.set(key, { data: row, rowNum: inputSheet.startRow + i });
      });

      // 1. 工数シートから削除すべき行を特定
      const rowsToDelete = [];
      for (const [key, value] of inputProjects.entries()) {
        if (!mainProjects.has(key)) {
          rowsToDelete.push(value.rowNum);
        }
      }

      // 2. 工数シートに追加すべき行を特定
      const rowsToAdd = [];
      for (const [key, value] of mainProjects.entries()) {
        if (!inputProjects.has(key)) {
          rowsToAdd.push([
            value[mainIndices.MGMT_NO - 1],
            value[mainIndices.SAGYOU_KUBUN - 1],
            value[mainIndices.KIBAN - 1],
            value[mainIndices.PROGRESS - 1] || "",
            value[mainIndices.PLANNED_HOURS - 1],
          ]);
        }
      }

      // 3. 差分更新を実行
      if (rowsToDelete.length > 0) {
        rowsToDelete.reverse().forEach(rowNum => inputSheet.sheet.deleteRow(rowNum));
      }
      if (rowsToAdd.length > 0) {
        const startWriteRow = inputSheet.getLastRow() + 1;
        inputSheet.sheet.getRange(startWriteRow, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
        
        const sumFormulas = [];
        const dateStartCol = Object.keys(INPUT_SHEET_HEADERS).length + 1;
        const dateStartColLetter = inputSheet.sheet.getRange(1, dateStartCol).getA1Notation().replace("1", "");
        for (let i = 0; i < rowsToAdd.length; i++) {
          const rowNum = startWriteRow + i;
          sumFormulas.push([`=IFERROR(SUM(${dateStartColLetter}${rowNum}:${rowNum}))`]);
        }
        inputSheet.sheet.getRange(startWriteRow, inputSheet.indices.ACTUAL_HOURS_SUM, rowsToAdd.length, 1).setFormulas(sumFormulas);
      }

    } catch (e) {
      Logger.log(`${tantousha.name} の工数シート差分更新中にエラー: ${e.message}`);
    }
  });
}

// =================================================================================
// === 工数 → メインシートへの同期処理 (変更なし) ===
// =================================================================================
function syncInputToMain(inputSheetName, editedRange) {
  const tantoushaName = inputSheetName.replace(CONFIG.SHEETS.INPUT_PREFIX, '');
  const inputSheet = new InputSheet(tantoushaName);
  const mainSheet = new MainSheet();
  const mainDataMap = mainSheet.getDataMap();

  const editedRow = editedRange.getRow();
  if (editedRow < inputSheet.startRow) return;

  const editedRowValues = inputSheet.sheet.getRange(editedRow, 1, 1, inputSheet.getLastColumn()).getValues()[0];
  const inputIndices = inputSheet.indices;
  const mainIndices = mainSheet.indices;
  const mgmtNo = editedRowValues[inputIndices.MGMT_NO - 1];
  const sagyouKubun = editedRowValues[inputIndices.SAGYOU_KUBUN - 1];
  const uniqueKey = `${mgmtNo}_${sagyouKubun}`;
  const targetRowInfo = mainDataMap.get(uniqueKey);
  if (!targetRowInfo) return;

  const targetRowNum = targetRowInfo.rowNum;
  const editedCol = editedRange.getColumn();
  const targetRange = mainSheet.sheet.getRange(targetRowNum, 1, 1, mainSheet.getLastColumn());
  const targetValues = targetRange.getValues()[0];
  const targetFormulas = targetRange.getFormulas()[0];

  const newRowData = targetValues.map((cellValue, i) => targetFormulas[i] || cellValue);
  
  if (editedCol === inputIndices.PROGRESS) {
    const newProgress = editedRowValues[inputIndices.PROGRESS - 1];
    newRowData[mainIndices.PROGRESS - 1] = newProgress;

    const completionTriggers = getCompletionTriggerStatuses();
    const startDateTriggers = getStartDateTriggerStatuses();
    if (!isValidDate(newRowData[mainIndices.START_DATE - 1]) && startDateTriggers.includes(newProgress)) {
      newRowData[mainIndices.START_DATE - 1] = new Date();
    }
    
    if (completionTriggers.includes(newProgress)) {
      newRowData[mainIndices.COMPLETE_DATE - 1] = new Date();
    } else {
      newRowData[mainIndices.COMPLETE_DATE - 1] = '';
    }
  }

  let totalHours = 0;
  const dateStartCol = Object.keys(INPUT_SHEET_HEADERS).length + 1;
  if (inputSheet.getLastColumn() >= dateStartCol) {
    const hoursValues = inputSheet.sheet.getRange(editedRow, dateStartCol, 1, inputSheet.getLastColumn() - dateStartCol + 1).getValues()[0];
    totalHours = hoursValues.reduce((sum, h) => sum + toNumber(h), 0);
  }
  newRowData[mainIndices.ACTUAL_HOURS - 1] = totalHours;
  newRowData[mainIndices.PROGRESS_EDITOR - 1] = tantoushaName;
  newRowData[mainIndices.UPDATE_TS - 1] = new Date();

  targetRange.setValues([newRowData]);
}