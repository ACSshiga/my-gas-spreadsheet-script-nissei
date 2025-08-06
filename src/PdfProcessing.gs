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

        const kishu = getValue(appText, /機種\s*[:：]\s*([\s\S]*?)(?=\s*机番\s*[:：]|\s*機番\s*[:：]|\s*納入先\s*[:：]|\s*・機械納期|\n)/, 2);
        const kiban = getValue(appText, /機