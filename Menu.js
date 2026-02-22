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
 * バックアップファイルのIDとシートデータをスクリプトキャッシュに格納する。
 * backupフォルダ内の "BACKUP_{スプレッドシート名}" を探す。
 * シートデータはバックアップファイルから読み込む（currentは既に編集済みの可能性があるため）。
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
  const cache = CacheService.getScriptCache();
  cache.put(CONFIG.CACHE.BACKUP_FILE_ID_KEY, backupId, CONFIG.CACHE.TTL_SECONDS);

  try {
    const backupSs = SpreadsheetApp.openById(backupId);
    backupSs.getSheets().forEach(sheet => {
      const data = sheet.getDataRange().getValues();
      const cacheKey = CONFIG.CACHE.BACKUP_DATA_PREFIX + sheet.getName();
      try {
        cache.put(cacheKey, JSON.stringify(data), CONFIG.CACHE.TTL_SECONDS);
      } catch (e) {
        console.warn(`[prepareBackupCache]: シート "${sheet.getName()}" のキャッシュ書き込みに失敗しました: ${e.toString()}`);
      }
    });
  } catch (e) {
    console.warn(`[prepareBackupCache]: バックアップのシートデータキャッシュに失敗しました: ${e.toString()}`);
  }

  console.log("[prepareBackupCache]: キャッシュを更新しました");
  return backupId;
}
