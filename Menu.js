/**
 * @fileoverview メニュー登録とバックアップキャッシュ準備
 * スプレッドシート起動時にカスタムメニューを追加し、
 * バックアップファイルのIDをスクリプトキャッシュにロードする。
 */

/**
 * スプレッドシートが開かれたときに実行されるシンプルトリガー。
 * ツールメニューを追加し、バックアップキャッシュを準備する。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠️ ツール')
    .addItem('Markdownエクスポート実行', 'exportSheetsToMarkdown')
    .addSeparator()
    .addItem('バックアップ更新（差分リセット）', 'resetAndBackup')
    .addToUi();

  prepareBackupCache();
}

/**
 * バックアップファイルのIDを特定してスクリプトキャッシュに格納する。
 * backupフォルダ内の "BACKUP_{スプレッドシート名}" を探す。
 * @returns {string|null} バックアップファイルのID。見つからない場合はnull
 */
function prepareBackupCache() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parentFolder = getParentFolder(ss.getId());
  const folders = parentFolder.getFoldersByName(CONFIG.FOLDER.BACKUP);

  if (!folders.hasNext()) return null;

  const backupFolder = folders.next();
  const files = backupFolder.getFilesByName(getBackupName(ss.getName()));

  if (!files.hasNext()) return null;

  const backupId = files.next().getId();
  CacheService.getScriptCache().put(
    CONFIG.CACHE.BACKUP_FILE_ID_KEY, backupId, CONFIG.CACHE.TTL_SECONDS
  );
  return backupId;
}
