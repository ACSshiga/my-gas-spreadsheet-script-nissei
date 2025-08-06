/**
 * PdfProcessing.gs
 * アップロードされた申請書ファイルを解析し、メインシートにデータをインポートする機能を担当します。
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
      
      Logger.log(`===== PDFファイル「${file.getName()}」から抽出したテキスト =====\n${text}\n========================================================`);

      const applications = text.split(/設計業務の外注委託申請書|--- PAGE \d+ ---/).filter(s => s.trim().length > 20 && /管理(N|Ｎ)(o|ｏ|O|Ｏ)(\.|．)/.test(s));
      
      if (applications.length === 0) {
        Logger.log(`ファイル「${file.getName()}」から有効な申請書データが見つかりませんでした。`);
        return;
      }

      applications.forEach((appText, i) => {
        Logger.log(`--- 申請書 ${i + 1} の解析開始 ---`);
        
        const cleanValue = (val) => val ? val.replace(/[\n\r\t]/g, ' ').replace(/\s+/g, ' ').trim() : '';
        
        // ★★★ここからが修正箇所★★★
        // グループのインデックスを5から4に修正
        const mgmtNoMatch = appText.match(/管理(N|Ｎ)(o|ｏ|O|Ｏ)(\.|．)\s*(\S+)/);
        const mgmtNo = mgmtNoMatch ? cleanValue(mgmtNoMatch[4]) : '';
        Logger.log(mgmtNo ? `[OK] 管理No: 「${mgmtNo}」` : `[NG] 管理Noが見つかりません。`);
        if (!mgmtNo) return;

        // 他の項目もより堅牢な正規表現に修正
        const kishuMatch = appText.match(/機種\s*[:：]([\s\S]*?)(?=\s*(机番|機番|納入先|・機械納期))/);
        const kishu = kishuMatch ? cleanValue(kishuMatch[1]) : '';
        Logger.log(kishu ? `[OK] 機種: 「${kishu}」` : `[INFO] 機種は空欄です。`);

        const kibanMatch = appText.match(/機番\s*[:：]([\s\S]*?)(?=\s*(納入先|・機械納期))/);
        const kiban = kibanMatch ? cleanValue(kibanMatch[1]) : '';
        Logger.log(kiban ? `[OK] 機番: 「${kiban}」` : `[INFO] 機番は空欄です。`);
        
        const nounyusakiMatch = appText.match(/納入先\s*[:：]([\s\S]*?)(?=\s*(・機械納期|入庫予定日))/);
        const nounyusaki = nounyusakiMatch ? cleanValue(nounyusakiMatch[1]) : '';
        Logger.log(nounyusaki ? `[OK] 納入先: 「${nounyusaki}」` : `[INFO] 納入先は空欄です。`);

        const kikanMatch = appText.replace(/\s/g, '').match(/設計予定期間[:：]?(\d{1,2}月\d{1,2}日)[~～-](\d{1,2}月\d{1,2}日)/);
        const sakuzuKigen = kikanMatch ? `${year}/${kikanMatch[2].replace('月', '/').replace('日', '')}` : '';
        Logger.log(sakuzuKigen ? `[OK] 作図期限: 「${sakuzuKigen}」` : `[INFO] 作図期限は見つかりませんでした。`);
        // ★★★ここまでが修正箇所★★★

        const kousuMatch = appText.match(/盤配\s*[:：]\s*(\d+)\s*H[\s\S]*?線加工\s*(\d+)\s*H/);
        if (kousuMatch) {
          const commonData = { mgmtNo, kishu, kiban, nounyusaki, sakuzuKigen };
          allNewRows.push(createRowData_(indices, { ...commonData, sagyouKubun: '盤配', yoteiKousu: cleanValue(kousuMatch[1]) }));
          allNewRows.push(createRowData_(indices, { ...commonData, sagyouKubun: '線加工', yoteiKousu: cleanValue(kousuMatch[2]) }));
        } else {
          const yoteiKousuMatch = appText.match(/見積設計工数\s*[:：]\s*(\d+)|(\d+)\s*Η|(\d+)\s*H/);
          const yoteiKousu = yoteiKousuMatch ? cleanValue(yoteiKousuMatch[1] || yoteiKousuMatch[2] || yoteiKousuMatch[3]) : '';
          
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
      syncMainToAllInputSheets();
      colorizeAllSheets();
    } else if (totalImportedCount > 0) {
      ui.alert(`${totalImportedCount}個のファイルを処理しましたが、シートに追加できる有効なデータが見つかりませんでした。Cloud Logsに詳細なデバッグ情報が出力されています。`);
    }

  } catch (e) {
    Logger.log(e.stack);
    ui.alert(`エラーが発生しました: ${e.message}`);
  }
}

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

function getValue(text, regex, groupIndex = 1) {
    const match = text.match(regex);
    return match && match[groupIndex] ? match[groupIndex].replace(/[\n\r\t]/g, ' ').replace(/\s+/g, ' ').trim() : '';
}

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