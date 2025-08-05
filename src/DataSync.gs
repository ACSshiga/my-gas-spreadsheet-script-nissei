/**
 * DataSync.gs
 * メインシートと工数シート間のデータ同期処理を専門に担当します。
 */

// =================================================================================
// === メイン → 工数シートへの同期処理 ===
// =================================================================================

/**
 * メインシートの変更を、全担当者の工数シートに反映させます。
 * 担当者のアサイン状況に応じて、各工数シートの行を動的に追加・削除します。
 */
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

      // メインシートから、その担当者の案件データのみをフィルタリング
      const tantoushaData = mainDataValues.filter(row => row[mainIndices.TANTOUSHA - 1] === tantousha.name);

      // ★★★ ここからが修正箇所 ★★★
      // 工数シートに表示するためのデータに変換
      const dataForInputSheet = tantoushaData.map(row => {
        return [
          row[mainIndices.MGMT_NO - 1],
          row[mainIndices.SAGYOU_KUBUN - 1],
          row[mainIndices.KIBAN - 1],
          // メインシートの進捗が空欄の場合に「未着手」を自動で設定する
          row[mainIndices.PROGRESS - 1] || "未着手", 
          row[mainIndices.PLANNED_HOURS - 1],
          // 実績工数合計列は数式なので空のまま
        ];
      });

      // 工数シートの既存データをクリアし、新しいデータで上書き
      inputSheet.clearData();
      if (dataForInputSheet.length > 0) {
        inputSheet.writeData(dataForInputSheet);
      }
    } catch (e) {
      // 担当者の工数シートが存在しない場合は何もしない
      Logger.log(`${tantousha.name} の工数シートが見つかりませんでした: ${e.message}`);
    }
  });
}

// =================================================================================
// === 工数 → メインシートへの同期処理 ===
// =================================================================================

/**
 * 特定の工数シートでの変更を検知し、メインシートに内容を反映させます。
 * @param {string} inputSheetName - 変更があった工数シートの名前
 * @param {GoogleAppsScript.Spreadsheet.Range} editedRange - 編集されたセル範囲
 */
function syncInputToMain(inputSheetName, editedRange) {
  const tantoushaName = inputSheetName.replace(CONFIG.SHEETS.INPUT_PREFIX, '');
  const inputSheet = new InputSheet(tantoushaName);
  const mainSheet = new MainSheet();
  const mainDataMap = mainSheet.getDataMap();

  const editedRow = editedRange.getRow();
  if (editedRow < inputSheet.startRow) return; // ヘッダーや合計行の編集は無視

  const editedRowValues = inputSheet.sheet.getRange(editedRow, 1, 1, inputSheet.getLastColumn()).getValues()[0];
  const inputIndices = inputSheet.indices;
  const mainIndices = mainSheet.indices;

  const mgmtNo = editedRowValues[inputIndices.MGMT_NO - 1];
  const sagyouKubun = editedRowValues[inputIndices.SAGYOU_KUBUN - 1];
  const uniqueKey = `${mgmtNo}_${sagyouKubun}`;
  const targetRowInfo = mainDataMap.get(uniqueKey);
  if (!targetRowInfo) return; // メインシートに対応する行がない場合は終了

  const valuesToUpdate = {};
  const editedCol = editedRange.getColumn();
  // 1. 進捗が変更された場合
  if (editedCol === inputIndices.PROGRESS) {
    const newProgress = editedRowValues[inputIndices.PROGRESS - 1];
    valuesToUpdate[mainIndices.PROGRESS] = newProgress;
    
    // 仕掛日と完了日の自動記録
    const mainRowData = targetRowInfo.data;
    const currentStartDate = mainRowData[mainIndices.START_DATE - 1];
    if (!isValidDate(currentStartDate) && newProgress !== '未着手') {
      valuesToUpdate[mainIndices.START_DATE] = new Date();
    }
    if (newProgress === '完了') {
       valuesToUpdate[mainIndices.COMPLETE_DATE] = new Date();
    } else {
       valuesToUpdate[mainIndices.COMPLETE_DATE] = ''; // 完了でなくなった場合は完了日をクリア
    }
  }

  // 2. 実績工数の集計 (進捗変更時も工数変更時も、常に再計算)
  let totalHours = 0;
  const dateStartCol = Object.keys(INPUT_SHEET_HEADERS).length + 1;
  if (inputSheet.getLastColumn() >= dateStartCol) {
    const hoursValues = inputSheet.sheet.getRange(editedRow, dateStartCol, 1, inputSheet.getLastColumn() - dateStartCol + 1).getValues()[0];
    totalHours = hoursValues.reduce((sum, h) => sum + toNumber(h), 0);
  }
  valuesToUpdate[mainIndices.ACTUAL_HOURS] = totalHours;
  // 3. 更新者と更新日時の記録
  valuesToUpdate[mainIndices.PROGRESS_EDITOR] = tantoushaName;
  valuesToUpdate[main.indices.UPDATE_TS] = new Date();
  // 4. メインシートの対応する行を一度に更新
  const updateRange = mainSheet.sheet.getRange(targetRowInfo.rowNum, 1, 1, mainSheet.getLastColumn());
  const newRowData = updateRange.getValues()[0];
  for (const [colIndex, value] of Object.entries(valuesToUpdate)) {
    newRowData[colIndex - 1] = value;
  }
  updateRange.setValues([newRowData]);
}