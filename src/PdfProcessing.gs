/**
 * DriveIntegration.gs
 * Google Driveとの連携機能（フォルダ作成、バックアップ）を管理します。
 */

// =================================================================================
// === Google Drive フォルダ作成関連機能 ===
// =================================================================================

/**
 * メインシートのデータに基づき、資料フォルダを一括作成し、シートにリンクを挿入します。
 * 機種(MODEL)については、指定された命名規則（先頭の英語+数字）に基づいてグルーピングされます。
 */
function bulkCreateMaterialFolders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const mainSheet = new MainSheet();
    const sheet = mainSheet.getSheet();
    const indices = mainSheet.indices;
    const requiredColumns = {
      KIBAN: indices.KIBAN, KIBAN_URL: indices.KIBAN_URL,
      MODEL: indices.MODEL, SERIES_URL: indices.SERIES_URL
    };
    for (const [key, value] of Object.entries(requiredColumns)) {
      if (!value) throw new Error(`必要な列「${key}」が見つかりません。`);
    }

    const lastRow = mainSheet.getLastRow();
    if (lastRow < mainSheet.startRow) return;
    const range = sheet.getRange(mainSheet.startRow, 1, lastRow - mainSheet.startRow + 1, mainSheet.getLastColumn());
    const values = range.getValues();
    const formulas = range.getFormulas();
    const processedItems = new Set();

    const folderCreationTasks = [
      {
        type: 'KIBAN', // 機番は完全一致
        valueCol: indices.KIBAN,
        linkCol: indices.KIBAN_URL,
        parentFolderId: CONFIG.FOLDERS.REFERENCE_MATERIAL_PARENT
      },
      {
        type: 'MODEL', // 機種は前方一致
        valueCol: indices.MODEL,
        linkCol: indices.SERIES_URL,
        parentFolderId: CONFIG.FOLDERS.SERIES_MODEL_PARENT
      }
    ];

    folderCreationTasks.forEach(task => {
      processedItems.clear();
      values.forEach((row, i) => {
        const originalValue = String(row[task.valueCol - 1]).trim();
        if (!originalValue) return;

        let groupValue;
        let folderName;

        if (task.type === 'MODEL') {
          const match = originalValue.match(/^[A-Za-z]+[0-9]+/);
          groupValue = match ? match[0] : originalValue; // マッチすればプレフィックス、しなければ元の値でグループ化
          folderName = groupValue; // フォルダ名はグループ名
        } else {
          groupValue = originalValue;
          folderName = originalValue;
        }

        if (groupValue && !processedItems.has(groupValue)) {
          const folderResult = getOrCreateFolder_(folderName, task.parentFolderId);
          if (folderResult && folderResult.folder) {
            const url = folderResult.folder.getUrl();
            if (task.type === 'MODEL') {
              // 新しい前方一致用の関数でリンクを更新
              updateLinksForModelPrefix_(values, formulas, groupValue, task.valueCol, task.linkCol, url);
            } else {
              // 既存の完全一致用の関数でリンクを更新
              updateLinksForSameValue_(values, formulas, groupValue, task.valueCol, task.linkCol, url);
            }
          }
          processedItems.add(groupValue);
        }
      });
    });

    const outputData = values.map((row, i) => row.map((cell, j) => formulas[i][j] || cell));
    range.setValues(outputData);

    ss.toast("資料フォルダの作成とリンク設定が完了しました。");
  } catch (error) {
    Logger.log(error.stack);
    ss.toast(`エラー: ${error.message}`);
  }
}

/**
 * (リファクタリング)
 * 指定された列の値が一致するすべての行に、ハイパーリンクを設定します。
 */
function updateLinksForSameValue_(allValues, allFormulas, valueToMatch, valueColumn, linkColumn, url) {
  allValues.forEach((row, i) => {
    if (String(row[valueColumn - 1]).trim() === valueToMatch) {
      if (!allFormulas[i][linkColumn - 1]) {
        // リンクの表示名は、マッチした値そのものを使用
        allFormulas[i][linkColumn - 1] = createHyperlinkFormula(url, valueToMatch);
      }
    }
  });
}

/**
 * (新規追加)
 * 指定されたプレフィックスに前方一致するすべての行に、ハイパーリンクを設定します。
 * @param {any[][]} allValues - シートの全データ値
 * @param {string[][]} allFormulas - シートの全数式
 * @param {string} prefixToMatch - 検索対象のプレフィックス
 * @param {number} valueColumn - プレフィックスを照合する列のインデックス
 * @param {number} linkColumn - リンクを挿入する列のインデックス
 * @param {string} url - 挿入するフォルダのURL
 */
function updateLinksForModelPrefix_(allValues, allFormulas, prefixToMatch, valueColumn, linkColumn, url) {
  allValues.forEach((row, i) => {
    const fullValue = String(row[valueColumn - 1]).trim();
    if (fullValue.startsWith(prefixToMatch)) {
      if (!allFormulas[i][linkColumn - 1]) {
        // リンクの表示名は、元の完全な機種名を使用
        allFormulas[i][linkColumn - 1] = createHyperlinkFormula(url, fullValue);
      }
    }
  });
}


/**
 * 指定された名前のフォルダを、指定された親フォルダ内に作成または取得する内部関数。
 */
function getOrCreateFolder_(name, parentFolderId) {
  // ... (この関数の内容は変更ありません)
}

// ... (週次バックアップ機能は変更ありません)