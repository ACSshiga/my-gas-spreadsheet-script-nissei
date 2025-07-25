/**
 * Nissei Spreadsheet Script
 * スプレッドシート操作用のGoogle Apps Script
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Nissei Tools')
    .addItem('データ処理実行', 'processData')
    .addItem('レポート生成', 'generateReport')
    .addToUi();
}

function processData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  console.log('データ処理を開始します...');
  
  // ここにスプレッドシート処理のロジックを追加
  sheet.getRange('A1').setValue('処理完了: ' + new Date());
}

function generateReport() {
  console.log('レポート生成中...');
  
  // レポート生成のロジックを追加
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange('B1').setValue('レポート生成: ' + new Date());
}