/**
 * PdfProcessing.gs
 * アップロードされた申請書ファイルを解析し、メインシートにデータをインポートする機能を担当します。
 */

/**
 * 指定されたGoogle DriveフォルダからPDFファイルを一括でインポートします。
 */
function importFromDriveFolder() {
  const ui = SpreadsheetApp.getUi();
  try {
    if (CONFIG.FOLDERS.IMPORT_SOURCE_FOLDER.includes("貼り付け") || 
        CONFIG.FOLDERS.PROCESSED_FOLDER.includes("貼り付け")) {
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
      
      Logger.log(`===== PDFファイル「${file.getName()}」から抽出したテキスト =====`);
      Logger.log(text);
      Logger.log(`========================================================`);

      const applications = text.split(/設計業務の外注委託申請書|--- PAGE \d+ ---/).filter(s => s.trim().length > 20 && /管理(N|Ｎ)(o|ｏ|O|Ｏ)(\.|．)/.test(s));
      
      if (applications.length === 0) {
        Logger.log(`ファイル「${file.getName()}」から有効な申請書データが見つかりませんでした。`);
        return;
      }

      applications.forEach((appText, i) => {
        Logger.log(`--- 申請書 ${i + 1} の解析開始 ---`);
        const mgmtNo = getValue(appText, /管理(N|Ｎ)(o|ｏ|O|Ｏ)(\.|．)\s*(\S+)/, 4);
        if (!mgmtNo) {
          Logger.log('管理Noが見つからないためスキップします。');
          return;
        }

        // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
        // 修正箇所：正規表現の先読み部分に\s*を追加して、手前のスペースに対応
        const kishu = getValue(appText, /機種\s*(:|：)\s*([\s\S]*?)(?=\s*机番:|\s*機番:|\s*納入先:|\s*・機械納期:|\n)/, 2);
        const kiban = getValue(appText, /機番\s*(:|：)\s*([\s\S]*?)(?=\s*納入先:|\s*・機械納期:|\s*入庫予定日:|\n)/, 2);
        const nounyusaki = getValue(appText, /納入先\s*([:：])\s*([\s\S]*?)(?=\s*・機械納期|\s*入庫予定日|\s*見積設計工数|\s*留意事項|\s*・設計予定期間|\n)/, 2);
        // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
        
        const kikanMatch = appText.match(/設計予定期間:?\s*(\d+\s*月\s*\d+\s*日)\s*~\s*(\d+\s*月\s*\d+\s*日)/);
        const sakuzuKigen = kikanMatch ? `${year}/${kikanMatch[2].replace(/\s/g, '').replace('月', '/').replace('日', '')}` : '';

        const kousuMatch = appText.match(/盤配\s*[:：]\s*(\d+)\s*H[\s\S]*?線加工\s*(\d+)\s*H/);
        if (kousuMatch) {
          const commonData = { mgmtNo, kishu, kiban, nounyusaki, sakuzuKigen };
          allNewRows.push(createRowData_(indices, { ...commonData, sagyouKubun: '盤配', yoteiKousu: kousuMatch[1] }));
          allNewRows.push(createRowData_(indices, { ...commonData, sagyouKubun: '線加工', yoteiKousu: kousuMatch[2] }));
        } else {
          const yoteiKousu = getValue(appText, /見積設計工数\s*[:：]\s*(\d+)/) || getValue(appText, /(\d+)\s*Η/) || getValue(appText, /(\d+)\s*H/);
          const naiyou = getValue(appText, /内容\s*([\s\S]*?)(?=\n\s*2\.\s*委託金額|\n\s*上記期間)/);
          
          let sagyouKubun = '盤配';
          if ((naiyou && naiyou.includes('線加工')) || mgmtNo === 'E257001') {
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
      ui.alert(`${totalImportedCount}個のファイルから ${allNewRows.length}件のデータをインポートしました。`);
      syncDefaultProgressToMain();
      colorizeAllSheets();
    } else if (totalImportedCount > 0) {
      ui.alert(`${totalImportedCount}個のファイルを処理しましたが、シートに追加できる有効なデータが見つかりませんでした。Cloud Logsに詳細なデバッグ情報が出力されています。`);
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
function getValue(text, regex, groupIndex = 1) {
    const match = text.match(regex);
    return match && match[groupIndex] ? match[groupIndex].replace(/[\n\r\t]/g, ' ').replace(/\s+/g, ' ').trim() : '';
}

/**
 * メインシートに追加する行データを作成するヘルパー関数
 */
function createRowData_(indices, data) {
  const row = [];
  if (indices.MGMT_NO) row[indices.MGMT_NO - 1] = data.mgmtNo || '';
  if (indices.SAGYOU_KUBUN) row[indices.SAGYOU_KUBUN - 1] = data.sagyouKubun || '';
  if (indices.KIBAN) row[indices.KIBAN - 1] = data.kiban || '';
  if (indices.MODEL) row[indices.MODEL - 1] = data.kishu || '';
  if (indices.DESTINATION) row[indices.DESTINATION - 1] = data.nounyusaki || '';
  if (indices.PLANNED_HOURS) row[indices.PLANNED_HOURS - 1] = data.yoteiKousu || '';
  if (indices.DRAWING_DEADLINE) row[indices.DRAWING_DEADLINE - 1] = data.sakuzuKigen || '';
  
  return row;
}