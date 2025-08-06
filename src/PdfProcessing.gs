/**
 * PdfProcessing.gs
 * アップロードされた申請書ファイルを解析し、メインシートにデータをインポートする機能を担当します。
 */

/**
 * ★★★ 新規追加 ★★★
 * 指定されたGoogle DriveフォルダからPDFファイルを一括でインポートします。
 */
function importFromDriveFolder() {
  try {
    const sourceFolder = DriveApp.getFolderById(CONFIG.FOLDERS.IMPORT_SOURCE_FOLDER);
    const processedFolder = DriveApp.getFolderById(CONFIG.FOLDERS.PROCESSED_FOLDER);
    const files = sourceFolder.getFilesByType(MimeType.PDF);
    
    let totalImportedCount = 0;
    const allNewRows = [];
    const mainSheet = new MainSheet();
    const indices = mainSheet.indices;
    const year = new Date().getFullYear();

    while (files.hasNext()) {
      const file = files.next();
      const text = extractTextFromPdf(file);
      
      const applications = text.split('設計業務の外注委託申請書').filter(Boolean);
      if (applications.length === 0) continue;

      applications.forEach(appText => {
        // (データ抽出ロジックは同じ)
        const mgmtNo = getValue(appText, /管理No\.\s*(\S+)/);
        const kishu = getValue(appText, /機種:\s*([^\s機]+)/);
        const kiban = getValue(appText, /機番:\s*(\S+)/);
        const nounyusaki = getValue(appText, /納入先:\s*(\S+)/);
        const kikanMatch = appText.match(/設計予定期間:\s*(\d+月\d+日)\s*~\s*(\d+月\d+日)/);
        const sakuzuKigen = kikanMatch ? `${year}/${kikanMatch[2].replace('月', '/').replace('日', '')}` : '';
        const kousuMatch = appText.match(/盤配:(\d+)H・線加工(\d+)H/);

        if (kousuMatch) {
          const commonData = { mgmtNo, kishu, kiban, nounyusaki, sakuzuKigen };
          allNewRows.push(createRowData_(indices, { ...commonData, sagyouKubun: '盤配', yoteiKousu: kousuMatch[1] }));
          allNewRows.push(createRowData_(indices, { ...commonData, sagyouKubun: '線加工', yoteiKousu: kousuMatch[2] }));
        } else {
          const yoteiKousu = getValue(appText, /見積設計工数:\s*(\d+)/);
          const naiyou = getValue(appText, /内容\s*([\s\S]*?)\n/);
          const sagyouKubun = (naiyou.includes('線加工')) ? '線加工' : '盤配';
          allNewRows.push(createRowData_(indices, { mgmtNo, sagyouKubun, kishu, kiban, nounyusaki, yoteiKousu, sakuzuKigen }));
        }
      });

      // 処理が終わったファイルを移動
      file.moveTo(processedFolder);
      totalImportedCount++;
    }

    if (allNewRows.length > 0) {
      const sheet = mainSheet.getSheet();
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, allNewRows.length, allNewRows[0].length).setValues(allNewRows);
      SpreadsheetApp.getActiveSpreadsheet().toast(`${totalImportedCount}個のファイルから ${allNewRows.length}件のデータをインポートしました。`);
      syncDefaultProgressToMain();
      colorizeAllSheets();
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast('インポート対象の新しいファイルはありませんでした。');
    }

  } catch (e) {
    Logger.log(e.stack);
    SpreadsheetApp.getUi().alert(`エラーが発生しました: ${e.message}`);
  }
}

/**
 * Drive上のPDFファイルからOCRでテキストを抽出します。
 * @param {File} file - Google Drive の File オブジェクト
 * @returns {string} 抽出したテキスト
 */
function extractTextFromPdf(file) {
  let tempDoc;
  try {
    const blob = file.getBlob();
    const resource = { title: `temp_ocr_${file.getName()}`, mimeType: MimeType.GOOGLE_DOCS };
    const tempDocFile = Drive.Files.insert(resource, blob, { ocr: true, ocrLanguage: 'ja' });
    tempDoc = DocumentApp.openById(tempDocFile.id);
    return tempDoc.getBody().getText();
  } catch(e) {
    throw new Error(`ファイル「${file.getName()}」のテキスト抽出に失敗しました: ${e.message}`);
  } finally {
    if (tempDoc) {
      Drive.Files.remove(tempDoc.getId());
    }
  }
}

// (getValue, createRowData_ のヘルパー関数は変更なし)
// ...