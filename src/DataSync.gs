/**
 * DataSync.gs
 * メインシートと工数シート間のデータ同期処理を専門に担当します。
 */

// =================================================================================
// === メイン → 工数シートへの同期処理 ===
// =================================================================================
function syncMainToAllInputSheets() {
  const mainSheet = new MainSheet();
  const mainDataValues = mainSheet.sheet.getRange(
    mainSheet.startRow, 1, 
    mainSheet.getLastRow() - mainSheet.startRow + 1, 
    mainSheet.getLastColumn()
  ).getValues();
  const mainIndices = mainSheet.indices;
  const tantoushaList = mainSheet.getTantoushaList();

  tantoushaList.forEach(tantousha => {
    try {
      const inputSheet = new InputSheet(tantousha.name);
      const tantoushaData = mainDataValues.filter(row => row[mainIndices.TANTOUSHA - 1] === tantousha.name);

      const dataForInputSheet = tantoushaData.map(row => [
        row[mainIndices.MGMT_NO - 1],
        row[mainIndices.SAGYOU_KUBUN - 1],
        row[mainIndices.KIBAN - 1],
        row[mainIndices.PROGRESS - 1] || "",
        row[mainIndices.PLANNED_HOURS - 1],
      ]);

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

  const updateRange = mainSheet.sheet.getRange(targetRowInfo.rowNum, 1, 1, mainSheet.getLastColumn());
  const newRowData = updateRange.getValues()[0];
  const editedCol = editedRange.getColumn();

  // 1. 進捗の更新
  if (editedCol === inputIndices.PROGRESS) {
    const newProgress = editedRowValues[inputIndices.PROGRESS - 1];
    newRowData[mainIndices.PROGRESS - 1] = newProgress;
    
    const currentStartDate = newRowData[mainIndices.START_DATE - 1];
    if (!isValidDate(currentStartDate) && newProgress !== '未着手') {
      newRowData[mainIndices.START_DATE - 1] = new Date();
    }
    newRowData[mainIndices.COMPLETE_DATE - 1] = (newProgress === '完了') ? new Date() : '';
  }

  // 2. 実績工数の再集計
  const dateStartCol = Object.keys(INPUT_SHEET_HEADERS).length + 1;
  let totalHours = 0;
  if (inputSheet.getLastColumn() >= dateStartCol) {
    const hoursValues = inputSheet.sheet.getRange(editedRow, dateStartCol, 1, inputSheet.getLastColumn() - dateStartCol + 1).getValues()[0];
    totalHours = hoursValues.reduce((sum, h) => sum + toNumber(h), 0);
  }
  newRowData[mainIndices.ACTUAL_HOURS - 1] = totalHours;

  // 3. 更新者と日時の記録
  newRowData[mainIndices.PROGRESS_EDITOR - 1] = tantoushaName;
  newRowData[mainSheet.indices.UPDATE_TS - 1] = new Date();

  // 4. メインシートへの一括書き込み
  updateRange.setValues([newRowData]);
}

/**
 * (リファクタリング)
 * メインシートで進捗が空の行に「未着手」をバッチ処理で設定します。
 */
function syncDefaultProgressToMain() {
  const mainSheet = new MainSheet();
  const lastRow = mainSheet.getLastRow();
  if (lastRow < mainSheet.startRow) return;

  const range = mainSheet.sheet.getRange(mainSheet.startRow, 1, lastRow - mainSheet.startRow + 1, mainSheet.getLastColumn());
  const values = range.getValues();
  const progressColIndex = mainSheet.indices.PROGRESS - 1;
  const tantoushaColIndex = mainSheet.indices.TANTOUSHA - 1;
  let hasUpdate = false;

  values.forEach(row => {
    const progress = row[progressColIndex];
    const tantousha = row[tantoushaColIndex];
    if (!progress && tantousha) {
      row[progressColIndex] = "未着手";
      hasUpdate = true;
    }
  });
  
  if (hasUpdate) {
    range.setValues(values);
  }
}