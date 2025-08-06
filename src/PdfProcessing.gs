/**
 * PdfProcessing.gs
 * アップロードされた申請書ファイルを解析し、メインシートにデータをインポートする機能を担当します。
 */

/**
 * 指定されたGoogle DriveフォルダからPDFファイルを一括でインポートします。
 */
function importFromDriveFolder() {
  const ui = SpreadsheetApp.getUi(); // ★★★ UI操作用の変数を最初に定義 ★★★
  try {
    if (CONFIG.FOLDERS.IMPORT_SOURCE_FOLDER === "ここに「申請書インポート用」のIDを貼り付け" || 
        CONFIG.FOLDERS.PROCESSED_FOLDER === "ここに「処理済み申請書」のIDを貼り付け") {
      ui.alert('エラー: Config.gsファイルにインポート用のフォルダIDが正しく設定されていません。');
      return;
    }

    const sourceFolder = DriveApp.getFolderById(CONFIG.FOLDERS.IMPORT_SOURCE_FOLDER);
    const processedFolder = DriveApp.getFolderById(CONFIG.FOLDERS.PROCESSED_FOLDER);
    const filesIterator = sourceFolder.getFilesByType(MimeType.PDF);
    
    const filesForProcessing = [];
    while (filesIterator.hasNext()) {
      filesForProcessing.push(filesIterator.next());
    }
    const fileCount = filesForProcessing.length;

    if (fileCount === 0) {
      ui.alert('インポート用フォルダ内にPDFファイルが見つかりませんでした。');
      return;
    }
    
    ui.alert(`インポート用フォルダ内で ${fileCount} 個のPDFファイルが見つかりました。処理を開始します。`);

    let totalImportedCount = 0;
    const allNewRows = [];
    const mainSheet = new MainSheet();
    const indices = mainSheet.indices;
    const year = new Date().getFullYear();

    filesForProcessing.forEach(file => {
      const text = extractTextFromPdf(file);
      
      const applications = text.split(/設計業務の外注委託申請書|--- PAGE \d+ ---/).filter(s => s.trim().length > 10);
      if (applications.length === 0) return;

      applications.forEach(appText => {
        const mgmtNo = getValue(appText, /管理No\.\s*(\S+)/);
        if (!mgmtNo) return;

        const kishu = getValue(appText, /機種:\s*(.*?)(?=\s*機番:|\n)/s);
        const kiban = getValue(appText, /機番:\s*(.*?)(?=\s*納入先:|\n)/s);
        const nounyusaki = getValue(appText, /納入先:\s*(.*?)(?=\n|・機械納期:)/s);
        
        const kikanMatch = appText.match(/設計予定期間:?\s*(\d+\s*月\s*\d+\s*日)\s*~\s*(\d+\s*月\s*\d+\s*日)/);
        const sakuzuKigen = kikanMatch ? `${year}/${kikanMatch[2].replace(/\s/g, '').replace('月', '/').replace('日', '')}` : '';

        const kousuMatch = appText.match(/盤配\s*:\s*(\d+)\s*H.*?線加工\s*(\d+)\s*H/s);

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
          if (mgmtNo === 'E257001') {
            sagyouKubun = '線加工';
          }
          
          allNewRows.push(createRowData_(indices, { mgmtNo, sagyouKubun, kishu, kiban, nounyusaki, yoteiKousu, sakuzuKigen }));
        }
      });

      file.moveTo(processedFolder);
      totalImportedCount++;
    });

    if (allNewRows.length > 0) {
      const sheet = mainSheet.getSheet();
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, allNewRows.length, allNewRows[0].length).setValues(allNewRows);
      // ★★★ ここから修正 ★★★
      ui.toast(`${totalImportedCount}個のファイルから ${allNewRows.length}件のデータをインポートしました。`);
      // ★★★ ここまで修正 ★★★
      syncDefaultProgressToMain();
      colorizeAllSheets();
    } else if (totalImportedCount > 0) {
      // ★★★ ここから修正 ★★★
      ui.toast(`${totalImportedCount}個のファイルを処理しましたが、シートに追加できる有効なデータが見つかりませんでした。`);
      // ★★★ ここまで修正 ★★★
    }

  } catch (e) {
    Logger.log(e.stack);
    ui.alert(`エラーが発生しました: ${e.message}`);
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
  return match && match[1] ? match[1].replace(/[\n\r\t]/g, ' ').trim() : '';
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