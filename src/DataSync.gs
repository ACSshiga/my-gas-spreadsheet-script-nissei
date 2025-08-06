/**
 * DataSync.gs
 * メインシートと工数シート間のデータ同期処理を専門に担当します。
 */

// =================================================================================
// === メイン → 工数シートへの同期処理 (ハイブリッドアプローチ版) ===
// =================================================================================
function syncMainToAllInputSheets() {
  const mainSheet = new MainSheet();
  const mainIndices = mainSheet.indices;
  const lastRow = mainSheet.getLastRow();

  // メインシートの「担当者」が割り当てられ、「進捗」が空の場合に「未着手」をセットする
  if (lastRow >= mainSheet.startRow) {
    const range = mainSheet.sheet.getRange(
      mainSheet.startRow, 1,
      lastRow - mainSheet.startRow + 1,
      mainSheet.getLastColumn()
    );
    const mainData = range.getValues();
    let hasUpdate = false;
    mainData.forEach(row => {
      const progress = row[mainIndices.PROGRESS - 1];
      const tantousha = row[mainIndices.TANTOUSHA - 1];
      if (!progress && tantousha) {
        row[mainIndices.PROGRESS - 1] = "未着手";
        hasUpdate = true;
      }
    });
    if (hasUpdate) {
      range.setValues(mainData);
    }
  }

  const tantoushaList = mainSheet.getTantoushaList();
  if (mainSheet.getLastRow() < mainSheet.startRow) {
    // メインシートが空なら全工数シートをクリア
    tantoushaList.forEach(tantousha => {
      try { (new InputSheet(tantousha.name)).clearData(); } catch (e) {}
    });
    return;
  }

  const mainDataValues = mainSheet.sheet.getRange(
    mainSheet.startRow, 1,
    mainSheet.getLastRow() - mainSheet.startRow + 1,
    mainSheet.getLastColumn()
  ).getValues();
  
  // 各担当者シートへの同期処理
  tantoushaList.forEach(tantousha => {
    try {
      const inputSheet = new InputSheet(tantousha.name);
      const existingProgressMap = inputSheet.getDataMapForProgress();
      const tantoushaData = mainDataValues.filter(row => row[mainIndices.TANTOUSHA - 1] === tantousha.name);
      
      const dataForInputSheet = tantoushaData.map(row => {
        const mgmtNo = row[mainIndices.MGMT_NO - 1];
        const sagyouKubun = row[mainIndices.SAGYOU_KUBUN - 1];
        const key = `${mgmtNo}_${sagyouKubun}`;
        
        // 工数シートに既存の進捗があればそれを使い、なければメインシートの進捗を使う
        const progressToSet = existingProgressMap.get(key) || row[mainIndices.PROGRESS - 1] || "";

        return [
          mgmtNo, sagyouKubun, row[mainIndices.KIBAN - 1],
          progressToSet, row[mainIndices.PLANNED_HOURS - 1],
        ];
      });

      inputSheet.clearData();
      if (dataForInputSheet.length > 0) {
        inputSheet.writeData(dataForInputSheet);
      }
    } catch (e) {
      Logger.log(`${tantousha.name} の工数シート処理中にエラー: ${e.message}`);
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

  const targetRowNum = targetRowInfo.rowNum;
  const editedCol = editedRange.getColumn();
  // ▼▼▼ 修正箇所 START: リンクが消えない＆値が正しく反映されるロジックに修正 ▼▼▼
  const targetRange = mainSheet.sheet.getRange(targetRowNum, 1, 1, mainSheet.getLastColumn());
  const targetValues = targetRange.getValues()[0];
  const targetFormulas = targetRange.getFormulas()[0];

  // 数式を保持した行データを作成
  const newRowData = targetValues.map((cellValue, i) => targetFormulas[i] || cellValue);
  // 1. 進捗の更新
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

  // 2. 実績工数の更新
  let totalHours = 0;
  const dateStartCol = Object.keys(INPUT_SHEET_HEADERS).length + 1;
  if (inputSheet.getLastColumn() >= dateStartCol) {
    const hoursValues = inputSheet.sheet.getRange(editedRow, dateStartCol, 1, inputSheet.getLastColumn() - dateStartCol + 1).getValues()[0];
    totalHours = hoursValues.reduce((sum, h) => sum + toNumber(h), 0);
  }
  newRowData[mainIndices.ACTUAL_HOURS - 1] = totalHours;
  // 3. 担当者と更新日時の更新
  newRowData[mainIndices.PROGRESS_EDITOR - 1] = tantoushaName;
  newRowData[mainIndices.UPDATE_TS - 1] = new Date();

  // 4. 修正した行データを一括で書き戻す
  targetRange.setValues([newRowData]);
  // ▲▲▲ 修正箇所 END ▲▲▲
}


/**
 * この関数は syncMainToAllInputSheets に統合されたため、現在は使用されません。
 */
function syncDefaultProgressToMain() {
  // 機能は syncMainToAllInputSheets に統合されました。
}