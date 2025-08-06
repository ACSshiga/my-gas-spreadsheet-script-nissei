/**
 * DataSync.gs
 * メインシートと工数シート間のデータ同期処理を専門に担当します。
 */

// =================================================================================
// === メイン → 工数シートへの同期処理 ===
// =================================================================================
function syncMainToAllInputSheets() {
  const mainSheet = new MainSheet();
  const tantoushaList = mainSheet.getTantoushaList();

  if (mainSheet.getLastRow() < mainSheet.startRow) {
    Logger.log("メインシートのデータが空のため、全工数シートをクリアします。");
    tantoushaList.forEach(tantousha => {
      try {
        const inputSheet = new InputSheet(tantousha.name);
        inputSheet.clearData();
      } catch (e) {
        Logger.log(`${tantousha.name} の工数シートが見つかりませんでした: ${e.message}`);
      }
    });
    return;
  }

  const mainDataValues = mainSheet.sheet.getRange(
    mainSheet.startRow, 1, 
    mainSheet.getLastRow() - mainSheet.startRow + 1, 
    mainSheet.getLastColumn()
  ).getValues();
  const mainIndices = mainSheet.indices;


  tantoushaList.forEach(tantousha => {
    try {
      const inputSheet = new InputSheet(tantousha.name);
      const tantoushaData = mainDataValues.filter(row => row[mainIndices.TANTOUSHA - 1] === tantousha.name);
      const dataForInputSheet = tantoushaData.map(row => {
        return [
          row[mainIndices.MGMT_NO - 1],
          row[mainIndices.SAGYOU_KUBUN - 1],
          row[mainIndices.KIBAN - 1],
          row[mainIndices.PROGRESS - 1] || "",
          row[mainIndices.PLANNED_HOURS - 1],
        ];
      });

      inputSheet.clearData();
      if (dataForInputSheet.length > 0) {
        inputSheet.writeData(dataForInputSheet);
      }
    } catch (e) {
      Logger.log(`${tantousha.name} の工数シートが見つかりませんでした: ${e.message}`);
    }
  });
}

// =================================================================================
// === 工数 → メインシートへの同期処理 ===
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

  const valuesToUpdate = {};
  const editedCol = editedRange.getColumn();
  if (editedCol === inputIndices.PROGRESS) {
    const newProgress = editedRowValues[inputIndices.PROGRESS - 1];
    valuesToUpdate[mainIndices.PROGRESS] = newProgress;
    // 仕掛日と完了日の自動記録
    const completionTriggers = getCompletionTriggerStatuses();
    const startDateTriggers = getStartDateTriggerStatuses();
    const mainRowData = targetRowInfo.data;
    const currentStartDate = mainRowData[mainIndices.START_DATE - 1];

    if (!isValidDate(currentStartDate) && startDateTriggers.includes(newProgress)) {
      valuesToUpdate[mainIndices.START_DATE] = new Date();
    }
    
    if (completionTriggers.includes(newProgress)) {
       valuesToUpdate[mainIndices.COMPLETE_DATE] = new Date();
    } else {
       valuesToUpdate[mainIndices.COMPLETE_DATE] = '';
    }
  }

  let totalHours = 0;
  const dateStartCol = Object.keys(INPUT_SHEET_HEADERS).length + 1;
  if (inputSheet.getLastColumn() >= dateStartCol) {
    const hoursValues = inputSheet.sheet.getRange(editedRow, dateStartCol, 1, inputSheet.getLastColumn() - dateStartCol + 1).getValues()[0];
    totalHours = hoursValues.reduce((sum, h) => sum + toNumber(h), 0);
  }
  valuesToUpdate[mainIndices.ACTUAL_HOURS] = totalHours;

  valuesToUpdate[mainIndices.PROGRESS_EDITOR] = tantoushaName;
  valuesToUpdate[mainSheet.indices.UPDATE_TS] = new Date();
  
  const updateRange = mainSheet.sheet.getRange(targetRowInfo.rowNum, 1, 1, mainSheet.getLastColumn());
  const newRowData = updateRange.getValues()[0];
  for (const [colIndex, value] of Object.entries(valuesToUpdate)) {
    newRowData[colIndex - 1] = value;
  }
  updateRange.setValues([newRowData]);
}

/**
 * 工数シートの「未着手」をメインシートに反映する関数
 */
function syncDefaultProgressToMain() {
  const mainSheet = new MainSheet();
  const lastRow = mainSheet.getLastRow();
  if (lastRow < mainSheet.startRow) return;
  
  const range = mainSheet.sheet.getRange(
    mainSheet.startRow, 1,
    lastRow - mainSheet.startRow + 1,
    mainSheet.getLastColumn()
  );
  const mainData = range.getValues();
  let hasUpdate = false;

  mainData.forEach(row => {
    const progress = row[mainSheet.indices.PROGRESS - 1];
    const tantousha = row[mainSheet.indices.TANTOUSHA - 1];
    
    if (!progress && tantousha) {
      row[mainSheet.indices.PROGRESS - 1] = "未着手";
      hasUpdate = true;
    }
  });
  if(hasUpdate) {
    range.setValues(mainData);
  }
}