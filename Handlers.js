/**
 * @fileoverview 編集イベントハンドラ
 * セル編集時にバックアップと比較し、変更行を青字でハイライトする。
 * インストール可能トリガー（onEdit）から呼び出される。
 */

/**
 * 編集時のメインハンドラ。単一セルの編集のみ処理する。
 * バックアップ値と比較し、差分がある行を青字にする。
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - 編集イベントオブジェクト
 */
function handleInstalledEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const newValue = range.getValue().toString();

  if (range.getNumRows() > 1 || range.getNumColumns() > 1) return;

  const backupValue = getBackupValueFast(sheetName, range);

  console.log(`Sheet: ${sheetName} | Cell: ${range.getA1Notation()}`);
  console.log(`Backup Value (first 20 chars): ${backupValue ? backupValue.substring(0, 20) : "NULL"}`);

  if (backupValue === null) return;

  applyLineDiffHighlight(range, newValue, backupValue);
}

/**
 * バックアップ値を高速取得する。
 * まずスクリプトキャッシュを参照し、ヒットしなければバックアップファイルから直接読む。
 * @param {string} sheetName - シート名
 * @param {GoogleAppsScript.Spreadsheet.Range} range - 対象セル範囲
 * @returns {string|null} バックアップ値の文字列。取得できない場合はnull
 */
function getBackupValueFast(sheetName, range) {
  const cache = CacheService.getScriptCache();
  const cacheKey = CONFIG.CACHE.BACKUP_DATA_PREFIX + sheetName;

  // 1. キャッシュから取得
  const cachedData = cache.get(cacheKey);
  if (cachedData) {
    try {
      const data = JSON.parse(cachedData);
      const row = range.getRow() - 1;
      const col = range.getColumn() - 1;
      if (data[row] && data[row][col] !== undefined) {
        return data[row][col].toString();
      }
    } catch (e) {
      console.error("Cache parse failed: " + e.toString());
    }
  }

  // 2. ファイルからフォールバック
  let backupId = cache.get(CONFIG.CACHE.BACKUP_FILE_ID_KEY);
  if (!backupId) backupId = prepareBackupCache();
  if (!backupId) return null;

  try {
    const backupFile = SpreadsheetApp.openById(backupId);
    const backupSheet = backupFile.getSheetByName(sheetName);
    if (!backupSheet) return null;
    return backupSheet.getRange(range.getA1Notation()).getValue().toString();
  } catch (err) {
    console.error("Backup file error: " + err);
    return null;
  }
}

/**
 * セル内改行（LF）ベースの行単位差分ハイライトを適用する。
 * バックアップに存在しない行を青字にし、変更なしなら既定色に戻す。
 * @param {GoogleAppsScript.Spreadsheet.Range} range - 対象セル範囲
 * @param {string} newValue - 編集後の値
 * @param {string} oldValue - バックアップの値
 */
function applyLineDiffHighlight(range, newValue, oldValue) {
  /** @param {string} v - 改行コードを正規化する対象文字列 */
  const normalize = (v) => v.toString().replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const cleanNew = normalize(newValue);
  const cleanOld = normalize(oldValue);

  const newLines = cleanNew.split("\n");
  const oldLinesSet = new Set(cleanOld.split("\n").map(line => line.trim()));

  const richTextBuilder = SpreadsheetApp.newRichTextValue().setText(cleanNew);
  const blueStyle = SpreadsheetApp.newTextStyle().setForegroundColor("blue").build();

  let hasChange = false;
  let currentPos = 0;

  newLines.forEach((line) => {
    const trimmedLine = line.trim();

    if (trimmedLine !== "" && !oldLinesSet.has(trimmedLine)) {
      const start = currentPos;
      const end = currentPos + line.length;

      if (start < cleanNew.length) {
        richTextBuilder.setTextStyle(start, Math.min(end, cleanNew.length), blueStyle);
        hasChange = true;
      }
    }
    currentPos += line.length + 1;
  });

  if (hasChange) {
    range.setRichTextValue(richTextBuilder.build());
  } else {
    const color = (range.getRow() === 1) ? "white" : "black";
    range.setFontColor(color);
  }
}
