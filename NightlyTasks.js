/**
 * @fileoverview ナイトリービルド（定時バッチ処理）
 * 時間主導型トリガーにより午前4時頃に実行される。
 * エクスポートとバックアップを順次実行する。
 */

/**
 * ナイトリービルドのエントリポイント。
 * エクスポートとバックアップを実行し、結果をログに記録する。
 * 時間主導型トリガーから呼び出される。
 */
function nightlyBuild() {
  const now = new Date();
  console.log(`[Nightly Build Start]: ${Utilities.formatDate(now, CONFIG.EXPORT.TIMEZONE, "yyyy-MM-dd HH:mm:ss")}`);
  try {
    exportSheetsToMarkdown();
    resetAndBackup();
    console.log("[Nightly Build Success]: Artifact generated.");
  } catch (e) {
    console.error("[Nightly Build Failed]: " + e.toString());
  }
}

/**
 * 差分ハイライトのリセットとバックアップの更新を行う。
 * 1. 全シートの文字色を初期状態（本文:黒、ヘッダー:白）にリセット
 * 2. 既存バックアップを削除し、現在のスプレッドシートをコピー
 * 3. バックアップIDとシートデータをスクリプトキャッシュに格納
 */
function resetAndBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cache = CacheService.getScriptCache();

  // 1. 全シートの文字色リセット
  ss.getSheets().forEach(sheet => {
    const lr = sheet.getLastRow(), lc = sheet.getLastColumn();
    if (lr > 0 && lc > 0) {
      sheet.getDataRange().setFontColor("black");
      sheet.getRange(1, 1, 1, lc).setFontColor("white");
    }
  });

  // 2. バックアップファイル作成
  const currentFile = DriveApp.getFileById(ss.getId());
  const parentFolder = getParentFolder(ss.getId());
  const backupFolder = getOrCreateSubFolder(parentFolder, CONFIG.FOLDER.BACKUP);
  const backupName = getBackupName(ss.getName());

  const oldFiles = backupFolder.getFilesByName(backupName);
  while (oldFiles.hasNext()) oldFiles.next().setTrashed(true);

  const backupFile = currentFile.makeCopy(backupName, backupFolder);

  // 3. IDとシートデータをキャッシュ
  cache.put(CONFIG.CACHE.BACKUP_FILE_ID_KEY, backupFile.getId(), CONFIG.CACHE.TTL_SECONDS);

  // シートデータ全体をキャッシュ（100KB/キー制限に注意: 大規模シートはキャッシュミスで直接読み込みにフォールバック）
  ss.getSheets().forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    const cacheKey = CONFIG.CACHE.BACKUP_DATA_PREFIX + sheet.getName();
    try {
      cache.put(cacheKey, JSON.stringify(data), CONFIG.CACHE.TTL_SECONDS);
    } catch (e) {
      console.warn(`Cache put failed for sheet "${sheet.getName()}": ${e.toString()}`);
    }
  });

  console.log("Backup and Multi-sheet Cache updated.");
}
