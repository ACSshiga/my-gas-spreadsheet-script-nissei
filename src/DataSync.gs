/**
 * DataSync.gs
 * メインシートと工数シート間のデータ同期処理を専門に担当します。
 */

// (syncMainToAllInputSheets 関数は変更なし)
// ...

// =================================================================================
// === 工数 → メインシートへの同期処理 ===
// =================================================================================
/**
 * 特定の工数シートでの変更を検知し、メインシートに内容を反映させます。
 */
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
    
    // ★★★ ここから修正 ★★★
    // 仕掛日と完了日の自動記録
    const completionTriggers = getCompletionTriggerStatuses(); // トリガーとなる進捗リストを取得
    const mainRowData = targetRowInfo.data;
    const currentStartDate = mainRowData[mainIndices.START_DATE - 1];

    if (!isValidDate(currentStartDate) && newProgress !== '未着手') {
      valuesToUpdate[mainIndices.START_DATE] = new Date();
    }
    
    // 新しい進捗がトリガーリストに含まれているかチェック
    if (completionTriggers.includes(newProgress)) {
       valuesToUpdate[mainIndices.COMPLETE_DATE] = new Date();
    } else {
       valuesToUpdate[mainIndices.COMPLETE_DATE] = ''; // トリガー以外は完了日をクリア
    }
    // ★★★ ここまで修正 ★★★
  }

  // 2. 実績工数の集計
  let totalHours = 0;
  const dateStartCol = Object.keys(INPUT_SHEET_HEADERS).length + 1;
  if (inputSheet.getLastColumn() >= dateStartCol) {
    const hoursValues = inputSheet.sheet.getRange(editedRow, dateStartCol, 1, inputSheet.getLastColumn() - dateStartCol + 1).getValues()[0];
    totalHours = hoursValues.reduce((sum, h) => sum + toNumber(h), 0);
  }
  valuesToUpdate[mainIndices.ACTUAL_HOURS] = totalHours;

  // 3. 更新者と更新日時の記録
  valuesToUpdate[mainIndices.PROGRESS_EDITOR] = tantoushaName;
  valuesToUpdate[mainSheet.indices.UPDATE_TS] = new Date();
  
  // 4. メインシートの対応する行を一度に更新
  const updateRange = mainSheet.sheet.getRange(targetRowInfo.rowNum, 1, 1, mainSheet.getLastColumn());
  const newRowData = updateRange.getValues()[0];
  for (const [colIndex, value] of Object.entries(valuesToUpdate)) {
    newRowData[colIndex - 1] = value;
  }
  updateRange.setValues([newRowData]);
}

// (syncDefaultProgressToMain 関数は変更なし)
// ...