/**
 * PdfProcessing.gs
 * アップロードされた申請書ファイルを解析し、メインシートにデータをインポートする機能を担当します。
 */

/**
 * アップロードされたPDFファイルをテキストに変換し、メインシートに行として追加します。
 * @param {object} fileData クライアントサイドから送られてきたファイル情報
 */
function importDataFromPdfFile(fileData) {
  let tempFile;
  try {
    // 一時的にファイルをGoogle Driveに作成してテキストを抽出
    const blob = Utilities.newBlob(fileData.bytes, fileData.mimeType, fileData.fileName);
    tempFile = Drive.Files.insert({ title: 'temp_pdf_import.pdf' }, blob, { ocr: true });
    const text = DocumentApp.openById(tempFile.id).getBody().getText();

    // テキスト抽出後、メインの処理関数を呼び出す
    importDataFromText(text);

  } catch (e) {
    Logger.log(e.stack);
    throw new Error(`PDFの解析中にエラーが発生しました: ${e.message}`);
  } finally {
    // 処理が終わったら一時ファイルを完全に削除
    if (tempFile) {
      Drive.Files.remove(tempFile.id);
    }
  }
}

/**
 * 抽出されたテキストデータを解析し、メインシートに行として追加します。
 * @param {string} text 抽出されたテキスト
 */
function importDataFromText(text) {
  const mainSheet = new MainSheet();
  const sheet = mainSheet.getSheet();
  const indices = main.getSheet().getIndices();

  const applications = text.split('設計業務の外注委託申請書').filter(Boolean);
  if (applications.length === 0) {
    throw new Error("申請書のデータが見つかりませんでした。");
  }

  const newRows = [];
  const year = new Date().getFullYear();

  applications.forEach(appText => {
    const mgmtNo = getValue(appText, /管理No\.\s*(\S+)/);
    const kishu = getValue(appText, /機種:\s*([^\s機]+)/);
    const kiban = getValue(appText, /機番:\s*(\S+)/);
    const nounyusaki = getValue(appText, /納入先:\s*(\S+)/);
    
    const kikanMatch = appText.match(/設計予定期間:\s*(\d+月\d+日)\s*~\s*(\d+月\d+日)/);
    const sakuzuKigen = kikanMatch ? `${year}/${kikanMatch[2].replace('月', '/').replace('日', '')}` : '';

    const kousuMatch = appText.match(/盤配:(\d+)H・線加工(\d+)H/);
    if (kousuMatch) {
      const commonData = { mgmtNo, kishu, kiban, nounyusaki, sakuzuKigen };
      newRows.push(createRowData_(indices, { ...commonData, sagyouKubun: '盤配', yoteiKousu: kousuMatch[1] }));
      newRows.push(createRowData_(indices, { ...commonData, sagyouKubun: '線加工', yoteiKousu: kousuMatch[2] }));
    } else {
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
    syncDefaultProgressToMain();
    colorizeAllSheets();
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('インポートできるデータがありませんでした。');
  }
}


function getValue(text, regex) {
  const match = text.match(regex);
  return match ? match[1].trim() : '';
}

function createRowData_(indices, data) {
  const row = [];
  row[indices.MGMT_NO - 1] = data.mgmtNo || '';
  row[indices.SAGYOU_KUBUN - 1] = data.sagyouKubun || '';
  row[indices.KIBAN - 1] = data.kiban || '';
  row[indices.MODEL - 1] = data.kishu || '';
  row[indices.DESTINATION - 1] = data.nounyusaki || '';
  row[indices.PLANNED_HOURS - 1] = data.yoteiKousu || '';
  row[indices.DRAWING_DEADLINE - 1] = data.sakuzuKigen || '';
  
  return row;
}