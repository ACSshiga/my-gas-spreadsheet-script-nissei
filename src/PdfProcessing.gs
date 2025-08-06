/**
 * PdfProcessing.gs
 * アップロードされた申請書ファイルを解析し、メインシートにデータをインポートする機能を担当します。
 */

/**
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
        const mgmtNo = getValue(appText, /管理No\.\s*(\S+)/);
        if (!mgmtNo) return; // 管理Noがなければスキップ

        // ★★★ ここから修正 ★★★
        const kishu = getValue(appText, /機種:\s*(.*?)(?=\s*機番:|\n)/);
        const kiban = getValue(appText, /機番:\s*(.*?)(?=\s*納入先:|\n)/);
        const nounyusaki = getValue(appText, /納入先:\s*(.*?)\n/);
        
        const kikanMatch = appText.match(/設計予定期間:\s*(\d+\s*月\s*\d+\s*日)\s*~\s*(\d+\s*月\s*\d+\s*日)/);
        const sakuzuKigen = kikanMatch ? `${year}/${kikanMatch[2].replace(/\s/g, '').replace('月', '/').replace('日', '')}` : '';

        // より柔軟な正規表現で、盤配と線加工の両方を検出
        const kousuMatch = appText.match(/盤配\s*:\s*(\d+)\s*H.*?線加工\s*(\d+)\s*H/);
        // ★★★ ここまで修正 ★★★

        if (kousuMatch) {
          const commonData = { mgmtNo, kishu, kiban, nounyusaki, sakuzuKigen };
          allNewRows.push(createRowData_(indices, { ...commonData, sagyouKubun: '盤配', yoteiKousu: kousuMatch[1] }));
          allNewRows.push(createRowData_(indices, { ...commonData, sagyouKubun: '線加工', yoteiKousu: kousuMatch[2] }));
        } else {
          const yoteiKousu = getValue(appText, /見積設計工数:\s*(\d+)/);
          const naiyou = getValue(appText, /内容\s*([\s\S]*?)(?=\n\s*2\.\s*委託金額|\n\s*上記期間)/);
          
          let sagyouKubun = '盤配';
          if (naiyou && naiyou.includes('線加工') && !naiyou.includes('盤配')) {
            sagyouKubun = '線加工';
          }
          if (mgmtNo === 'E257001') { // E257001の特殊ケースに対応
            sagyouKubun = '線加工';
          }
          
          allNewRows.push(createRowData_(indices, { mgmtNo, sagyouKubun, kishu, kiban, nounyusaki, yoteiKousu, sakuzuKigen }));
        }
      });

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
 */
function extractTextFromPdf(file) {
  let tempDoc;
  try {
    const blob = file.getBlob();
    const resource = { title: `temp_ocr_${file.getName()}` };
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


/**
 * テキストから正規表現で値を抽出するヘルパー関数
 */
function getValue(text, regex) {
  const match = text.match(regex);
  return match ? match[1].trim() : '';
}

/**
 * メインシートに追加する行データを作成するヘルパー関数
 */
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