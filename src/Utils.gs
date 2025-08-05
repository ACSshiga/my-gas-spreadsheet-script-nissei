/**
 * Utils.gs
 * 汎用ユーティリティ関数を定義します。
 * システム全体で使用される共通の補助関数です。
 */

// =================================================================================
// === 日付関連ユーティリティ ===
// =================================================================================
// ... (既存のコードはそのまま) ...

function getTantoushaNameByEmail(email) {
  if (!email) return null;
  const userEmail = email.trim();
  const masterData = getMasterData(CONFIG.SHEETS.TANTOUSHA_MASTER, 2);
  const user = masterData.find(row => String(row[1]).trim() === userEmail);
  return user ? user[0] : null;
}

function logWithTimestamp(message) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), DATE_FORMATS.DATETIME);
  console.log(`[${timestamp}] ${message}`);
}

// =================================================================================
// === ★★★ 新規追加 ★★★ ===
// =================================================================================
/**
 * スクリプトが使用するすべてのキャッシュを削除します。
 * カスタムメニューから実行できます。
 */
function clearScriptCache() {
  try {
    CacheService.getScriptCache().removeAll();
    SpreadsheetApp.getActiveSpreadsheet().toast('スクリプトのキャッシュをクリアしました。', '完了', 3);
    logWithTimestamp("スクリプトのキャッシュがクリアされました。");
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`キャッシュのクリア中にエラーが発生しました: ${e.message}`, 'エラー', 5);
    Logger.log(`キャッシュのクリア中にエラー: ${e.message}`);
  }
}