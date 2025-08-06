/**
 * Billing.gs
 * 請求シートの作成と更新に関する機能を担当します。
 */

// =================================================================================
// === 請求シートのメイン機能 ===
// =================================================================================

/**
 * カスタムメニューから呼び出され、請求シートを最新の状態に更新します。
 * 指定された月の「完了日」を持つ案件をメインシートから抽出し、請求シートに転記します。
 * @param {string} selectedMonth - 'YYYY-MM'形式の月の文字列 (例: '2025-08')
 */
function updateBillingSheet(selectedMonth) {
  if (!selectedMonth) {
    SpreadsheetApp.getUi().alert('月が選択されていません。');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = new MainSheet();
    const mainIndices = mainSheet.indices;

    // メインシートの全データを取得
    const mainData = mainSheet.sheet.getRange(
      mainSheet.startRow, 1,
      mainSheet.getLastRow() - mainSheet.startRow + 1,
      mainSheet.getLastColumn()
    ).getValues();

    // 選択された月に完了した案件のみをフィルタリング
    const [year, month] = selectedMonth.split('-').map(Number);
    const billingData = mainData.filter(row => {
      const completionDate = row[mainIndices.COMPLETE_DATE - 1];
      if (isValidDate(completionDate)) {
        return completionDate.getFullYear() === year && completionDate.getMonth() + 1 === month;
      }
      return false;
    });

    // 請求シートに書き出すためのデータに整形
    const dataForBillingSheet = billingData.map(row => {
      return [
        row[mainIndices.MGMT_NO - 1],       // 管理No.
        row[mainIndices.KIBAN - 1],          // 委託業務内容 (機番)
        row[mainIndices.SAGYOU_KUBUN - 1],   // 作業区分
        row[mainIndices.PLANNED_HOURS - 1],  // 予定工数
        row[mainIndices.ACTUAL_HOURS - 1]    // 実工数
      ];
    });
    
    // 請求シートを取得（なければ作成）
    let billingSheet = ss.getSheetByName(CONFIG.SHEETS.BILLING);
    if (!billingSheet) {
      billingSheet = ss.insertSheet(CONFIG.SHEETS.BILLING);
    }

    // シートの内容をクリアし、ヘッダーと新しいデータを書き込み
    billingSheet.clear();
    const headers = ["管理No", "委託業務内容", "作業区分", "予定工数", "実工数"];
    billingSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");

    if (dataForBillingSheet.length > 0) {
      billingSheet.getRange(2, 1, dataForBillingSheet.length, dataForBillingSheet[0].length).setValues(dataForBillingSheet);
    }
    
    billingSheet.autoResizeColumns(1, headers.length);
    ss.toast(`${year}年${month}月 分の請求シートを更新しました。`, '完了', 5);
    billingSheet.activate();
  } catch (error) {
    Logger.log(`請求シートの更新中にエラーが発生しました: ${error.stack}`);
    SpreadsheetApp.getUi().alert(`エラー: ${error.message}`);
  }
}

/**
 * 請求用のサイドバーを表示します。
 */
function showBillingSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('BillingSidebar')
      .setTitle('請求月を選択')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * サイドバーに表示する月のリストを取得します。
 * メインシートの「完了日」からユニークな年月を抽出します。
 * @return {{value: string, text: string}[]}
 */
function getBillingMonthOptions() {
  const mainSheet = new MainSheet();
  const mainIndices = mainSheet.indices;
  if (!mainIndices.COMPLETE_DATE) return [];

  const mainData = mainSheet.sheet.getRange(
    mainSheet.startRow, mainIndices.COMPLETE_DATE,
    mainSheet.getLastRow() - mainSheet.startRow + 1, 1
  ).getValues();

  const monthSet = new Set();
  mainData.forEach(row => {
    const date = row[0];
    if (isValidDate(date)) {
      const year = date.getFullYear();
      const month = date.getMonth() + 1;
      monthSet.add(`${year}-${String(month).padStart(2, '0')}`); // YYYY-MM 形式でセットに追加
    }
  });

  // セットから配列に変換し、降順にソート
  return Array.from(monthSet).sort().reverse().map(monthStr => {
      const [year, month] = monthStr.split('-');
      return {
          value: monthStr, // 'YYYY-MM'
          text: `${year}年${Number(month)}月` // 'YYYY年M月'
      };
  });
}