/**
 * DataSync.gs
 * メインシートと工数シート間のデータ同期処理を専門に担当します。
 */

// =================================================================================
// === メイン → 工数シートへの同期処理 ===
// =================================================================================
function syncMainToAllInputSheets() {
  const mainSheet = new MainSheet();
  const mainIndices = mainSheet.indices;
  const lastRow = mainSheet.getLastRow();

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
    if(hasUpdate) {
      range.setValues(mainData);
    }
  }
  
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
// === 工数 → メインシートへの同期処理 (診断ログ強化版) ===
// =================================================================================
function syncInputToMain(inputSheetName, editedRange) {
  Logger.log(`\n\n===== 同期処理開始 (syncInputToMain) =====`);
  Logger.log(`編集されたシート: ${inputSheetName}, 編集されたセル: ${editedRange.getA1Notation()}`);

  const tantoushaName = inputSheetName.replace(CONFIG.SHEETS.INPUT_PREFIX, '');
  const inputSheet = new InputSheet(tantoushaName);
  const mainSheet = new MainSheet();
  const mainDataMap = mainSheet.getDataMap();

  const editedRow = editedRange.getRow();
  if (editedRow < inputSheet.startRow) {
    Logger.log(`編集された行(${editedRow}行目)はヘッダー領域のため処理を終了します。`);
    Logger.log(`========================================`);
    return;
  }

  Logger.log(`メインシートのデータマップから ${mainDataMap.size} 件のキーを読み込みました。`);

  const editedRowValues = inputSheet.sheet.getRange(editedRow, 1, 1, inputSheet.getLastColumn()).getValues()[0];
  const inputIndices = inputSheet.indices;
  const mainIndices = mainSheet.indices;

  const mgmtNo = editedRowValues[inputIndices.MGMT_NO - 1];
  const sagyouKubun = editedRowValues[inputIndices.SAGYOU_KUBUN - 1];

  Logger.log(`工数シートの ${editedRow} 行目からキーを読み取りました:`);
  Logger.log(`  - 管理No: 「${mgmtNo}」 (型: ${typeof mgmtNo})`);
  Logger.log(`  - 作業区分: 「${sagyouKubun}」 (型: ${typeof sagyouKubun})`);

  if (!mgmtNo || !sagyouKubun) {
    Logger.log(`キー（管理Noまたは作業区分）が空のため、処理を中断します。`);
    Logger.log(`========================================`);
    return;
  }

  const uniqueKey = `${mgmtNo}_${sagyouKubun}`;
  Logger.log(`作成された検索キー: 「${uniqueKey}」`);

  const targetRowInfo = mainDataMap.get(uniqueKey);

  if (!targetRowInfo) {
    Logger.log(`[エラー!] 作成された検索キー「${uniqueKey}」がメインシートのデータマップに見つかりませんでした。`);
    Logger.log(`メインシートにこの「管理No」と「作業区分」の組み合わせが存在するか、完全一致（スペース等含む）しているか確認してください。`);
    Logger.log(`===== 同期処理終了 (対象が見つからず) =====`);
    return;
  }

  Logger.log(`[成功] メインシートで一致する行を発見しました。行番号: ${targetRowInfo.rowNum}`);

  const targetRowNum = targetRowInfo.rowNum;
  const editedCol = editedRange.getColumn();

  const targetRange = mainSheet.sheet.getRange(targetRowNum, 1, 1, mainSheet.getLastColumn());
  const targetValues = targetRange.getValues()[0];
  const targetFormulas = targetRange.getFormulas()[0];
  const newRowData = targetValues.map((cellValue, i) => targetFormulas[i] || cellValue);

  Logger.log(`メインシートの ${targetRowNum} 行目の更新処理を開始します...`);

  if (editedCol === inputIndices.PROGRESS) {
    const newProgress = editedRowValues[inputIndices.PROGRESS - 1];
    newRowData[mainIndices.PROGRESS - 1] = newProgress;
    Logger.log(`  -> 進捗を「${newProgress}」に更新します。`);

    const completionTriggers = getCompletionTriggerStatuses();
    const startDateTriggers = getStartDateTriggerStatuses();

    if (!isValidDate(newRowData[mainIndices.START_DATE - 1]) && startDateTriggers.includes(newProgress)) {
      newRowData[mainIndices.START_DATE - 1] = new Date();
      Logger.log(`  -> 仕掛日を自動入力しました。`);
    }

    if (completionTriggers.includes(newProgress)) {
      newRowData[mainIndices.COMPLETE_DATE - 1] = new Date();
      Logger.log(`  -> 完了日を自動入力しました。`);
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
  Logger.log(`  -> 実績工数を「${totalHours}」に更新します。`);

  newRowData[mainIndices.PROGRESS_EDITOR - 1] = tantoushaName;
  newRowData[mainIndices.UPDATE_TS - 1] = new Date();

  targetRange.setValues([newRowData]);
  Logger.log(`メインシートへの書き込みが完了しました。`);
  Logger.log(`===== 同期処理正常終了 =====`);
}