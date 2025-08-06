/**
 * PdfProcessing.gs
 * 貼り付けられた申請書のテキストを解析し、メインシートにデータをインポートする機能を担当します。
 */

/**
 * 貼り付けられたテキストデータを解析し、メインシートに行として追加します。
 * @param {string} text ユーザーがダイアログに貼り付けたテキスト
 */
function importDataFromPdfText(text) {
  try {
    const mainSheet = new MainSheet();
    const sheet = mainSheet.getSheet();
    const indices = mainSheet.indices;

    // 正規表現で各申請書を分割
    const applications = text.split('設計業務の外注委託申請書').filter(Boolean);
    if (applications.length === 0) {
      throw new Error("申請書のデータが見つかりませんでした。");
    }

    const newRows = [];
    let year = new Date().getFullYear(); // 年をまたぐ場合などは手動修正を想定

    applications.forEach(appText => {
      const mgmtNo = getValue(appText, /管理No\.\s*(\S+)/);
      const kishu = getValue(appText, /機種:\s*(\S+)/);
      const kiban = getValue(appText, /機番:\s*(\S+)/);
      const nounyusaki = getValue(appText, /納入先:\s*(\S+)/);
      
      // 設計予定期間から作図期限を抽出
      const kikanMatch = appText.match(/設計予定期間:\s*(\d+月\d+日)\s*~\s*(\d+月\d+日)/);
      const sakuzuKigen = kikanMatch ? `${year}/${kikanMatch[2].replace('月', '/').replace('日', '')}` : '';

      // 予定工数の解析
      const kousuMatch = appText.match(/盤配:(\d+)H・線加工(\d+)H/);
      if (kousuMatch) {
        // 盤配と線加工が両方ある場合
        const commonData = { mgmtNo, kishu, kiban, nounyusaki, sakuzuKigen };
        newRows.push(createRowData_(indices, { ...commonData, sagyouKubun: '盤配', yoteiKousu: kousuMatch[1] }));
        newRows.push(createRowData_(indices, { ...commonData, sagyouKubun: '線加工', yoteiKousu: kousuMatch[2] }));
      } else {
        // 単独の場合
        const yoteiKousu = getValue(appText, /見積設計工数:\s*(\d+)/);
        const naiyou = getValue(appText, /内容\s*([\s\S]*?)\n/);
        const sagyouKubun = (naiyou.includes('線加工')) ? '線加工' : '盤配';
        
        newRows.push(createRowData_(indices, { mgmtNo, sagyouKubun, kishu, kiban, nounyusaki, yoteiKousu, sakuzuKigen }));
      }
    });

    if (newRows.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      SpreadsheetApp.getActiveSpreadsheet().toast(`${newRows.length} 件のデータをインポートしました。`);
      // データの同期と色付けを実行
      syncDefaultProgressToMain();
      colorizeAllSheets();
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast('インポートできるデータがありませんでした。');
    }

  } catch (e) {
    Logger.log(e.stack);
    SpreadsheetApp.getUi().alert(`エラーが発生しました: ${e.message}`);
  }
}

/**
 * テキストから正規表現で値を抽出するヘルパー関数
 * @param {string} text - 解析対象のテキスト
 * @param {RegExp} regex - 抽出用の正規表現
 * @returns {string} 抽出した値、見つからなければ空文字
 */
function getValue(text, regex) {
  const match = text.match(regex);
  return match ? match[1].trim() : '';
}

/**
 * メインシートに追加する行データを作成するヘルパー関数
 * @param {object} indices - 列のインデックス情報
 * @param {object} data - 抽出したデータ
 * @returns {Array} シートに書き込むための1行分の配列
 */
function createRowData_(indices, data) {
  const row = [];
  // MAIN_SHEET_HEADERS の順番に合わせてデータを格納
  row[indices.MGMT_NO - 1] = data.mgmtNo || '';
  row[indices.SAGYOU_KUBUN - 1] = data.sagyouKubun || '';
  row[indices.KIBAN - 1] = data.kiban || '';
  row[indices.MODEL - 1] = data.kishu || '';
  row[indices.DESTINATION - 1] = data.nounyusaki || '';
  row[indices.PLANNED_HOURS - 1] = data.yoteiKousu || '';
  row[indices.DRAWING_DEADLINE - 1] = data.sakuzuKigen || '';
  
  return row;
}