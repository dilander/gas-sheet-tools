/**
 * @fileoverview 共通定数・ヘルパー関数
 * 全スクリプトから参照される設定値とユーティリティを集約する。
 */

/** @const {Object} アプリケーション共通設定 */
const CONFIG = {
  /** @property {Object} FOLDER - フォルダ名定義 */
  FOLDER: {
    BACKUP: "backup",
    EXPORT: "export",
  },
  /** @property {Object} CACHE - キャッシュ関連定義 */
  CACHE: {
    BACKUP_FILE_ID_KEY: "BACKUP_FILE_ID",
    BACKUP_DATA_PREFIX: "backup_data_",
    /** @type {number} キャッシュTTL（秒）6時間 */
    TTL_SECONDS: 21600,
  },
  /** @property {Object} EXPORT - エクスポート設定 */
  EXPORT: {
    EXTENSION: ".txt",
    DATE_FORMAT: "yyyyMMdd_HHmmss",
    TIMEZONE: "JST",
  },
  /** @property {Object} BACKUP - バックアップ設定 */
  BACKUP: {
    NAME_PREFIX: "BACKUP_",
    /** @type {string} キャッシュに載せるシート範囲（100KB制限に注意） */
    CACHE_RANGE: "A1:C100",
  },
  /** @type {string} エクスポート除外列の識別キーワード */
  IGNORE_KEYWORD: "無視",
  /** @type {number} Markdown見出しのオフセット基準レベル */
  HEADING_BASE_LEVEL: 4,
  /** @type {RegExp} コードブロック化する列名パターン */
  CODE_LANGUAGES: /(Python|Handlebars|Mermaid)/i,
};

/**
 * アクティブなスプレッドシートの親フォルダを取得する。
 * @param {string} ssId - スプレッドシートのファイルID
 * @returns {GoogleAppsScript.Drive.Folder} 親フォルダ
 * @throws {Error} 親フォルダが存在しない場合
 */
function getParentFolder(ssId) {
  const parents = DriveApp.getFileById(ssId).getParents();
  if (!parents.hasNext()) throw new Error("親フォルダが見つかりません。");
  return parents.next();
}

/**
 * 親フォルダ内のサブフォルダを取得する。存在しなければ新規作成する。
 * @param {GoogleAppsScript.Drive.Folder} parentFolder - 親フォルダ
 * @param {string} folderName - サブフォルダ名
 * @returns {GoogleAppsScript.Drive.Folder} サブフォルダ
 */
function getOrCreateSubFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}

/**
 * バックアップファイル名を生成する。
 * @param {string} ssName - スプレッドシート名
 * @returns {string} "BACKUP_" + スプレッドシート名
 */
function getBackupName(ssName) {
  return CONFIG.BACKUP.NAME_PREFIX + ssName;
}

/**
 * 現在日時のフォーマット済み文字列を返す。
 * @returns {string} "yyyyMMdd_HHmmss" 形式のタイムスタンプ
 */
function getTimestamp() {
  return Utilities.formatDate(new Date(), CONFIG.EXPORT.TIMEZONE, CONFIG.EXPORT.DATE_FORMAT);
}
