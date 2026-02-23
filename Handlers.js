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
      console.error(`[getBackupValueFast]: キャッシュのパースに失敗しました: ${e.toString()}`);
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
  } catch (e) {
    console.error(`[getBackupValueFast]: バックアップファイルの読み込みに失敗しました: ${e.toString()}`);
    return null;
  }
}

/**
 * セル内改行（LF）ベースの行単位差分ハイライトを適用する。
 * 変更がある場合、セル全体の背景色を設定し、変更文字に色を付ける。
 * 1行目（ヘッダー）は別の色でハイライトする。
 * 変更なしなら既定色と背景に戻す。
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
  const oldLines = cleanOld.split("\n");

  const isHeader = range.getRow() === 1;
  const textColor = isHeader
    ? CONFIG.DIFF_HIGHLIGHT.HEADER_TEXT_COLOR
    : CONFIG.DIFF_HIGHLIGHT.TEXT_COLOR;
  const bgColor = isHeader
    ? CONFIG.DIFF_HIGHLIGHT.HEADER_BACKGROUND_COLOR
    : CONFIG.DIFF_HIGHLIGHT.BACKGROUND_COLOR;

  const richTextBuilder = SpreadsheetApp.newRichTextValue().setText(cleanNew);
  const diffTextStyle = SpreadsheetApp.newTextStyle().setForegroundColor(textColor).build();

  let hasChange = cleanNew !== cleanOld;
  let currentPos = 0;

  newLines.forEach((newLine, lineIdx) => {
    const oldLine = oldLines[lineIdx] || "";

    // 行単位で文字単位の差分を検出
    for (let i = 0; i < newLine.length; i++) {
      if (newLine[i] !== oldLine[i]) {
        const charPos = currentPos + i;
        if (charPos < cleanNew.length) {
          richTextBuilder.setTextStyle(charPos, charPos + 1, diffTextStyle);
        }
      }
    }

    currentPos += newLine.length + 1;
  });

  if (hasChange) {
    range.setRichTextValue(richTextBuilder.build());
    range.setBackground(bgColor);
  } else {
    const color = isHeader ? CONFIG.CELL_COLORS.HEADER_TEXT : CONFIG.CELL_COLORS.DEFAULT_TEXT;
    const background = isHeader ? CONFIG.CELL_COLORS.HEADER_BACKGROUND : CONFIG.CELL_COLORS.DEFAULT_BACKGROUND;
    range.setFontColor(color);
    range.setBackground(background);
  }
}
