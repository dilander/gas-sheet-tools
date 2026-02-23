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
 * 行レベルのLCS（最長共通部分列）を計算し、一致行のペアを返す。
 * @param {string[]} oldLines - 変更前の行配列
 * @param {string[]} newLines - 変更後の行配列
 * @returns {Map<number, number>} newIndex → oldIndex のマッピング（完全一致行のみ）
 */
function computeLineLCS(oldLines, newLines) {
  const m = oldLines.length;
  const n = newLines.length;
  const dp = [];
  for (let i = 0; i <= m; i++) {
    dp[i] = new Array(n + 1).fill(0);
  }

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (oldLines[i - 1] === newLines[j - 1]) {
        dp[i][j] = dp[i - 1][j - 1] + 1;
      } else {
        dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
      }
    }
  }

  const matchMap = new Map();
  let i = m;
  let j = n;
  while (i > 0 && j > 0) {
    if (oldLines[i - 1] === newLines[j - 1]) {
      matchMap.set(j - 1, i - 1);
      i--;
      j--;
    } else if (dp[i - 1][j] > dp[i][j - 1]) {
      i--;
    } else {
      j--;
    }
  }

  return matchMap;
}

/**
 * 各新規行に対応する旧行インデックスを返す。
 * LCSで完全一致行をアンカーとし、アンカー間のギャップ内で位置ベースのペアリングを行う。
 * @param {string[]} oldLines - 変更前の行配列
 * @param {string[]} newLines - 変更後の行配列
 * @returns {number[]} 各newLineに対応するoldLineのインデックス。対応なしは-1
 */
function buildOldLineMapping(oldLines, newLines) {
  const matchMap = computeLineLCS(oldLines, newLines);
  const mapping = new Array(newLines.length).fill(-1);

  for (const [newIdx, oldIdx] of matchMap) {
    mapping[newIdx] = oldIdx;
  }

  // アンカーポイントをソートして取得
  const anchors = Array.from(matchMap.entries()).sort((a, b) => a[0] - b[0]);
  const sentineled = [[-1, -1], ...anchors, [newLines.length, oldLines.length]];

  // アンカー間のギャップで位置ベースのペアリング
  for (let a = 0; a < sentineled.length - 1; a++) {
    const gapNewStart = sentineled[a][0] + 1;
    const gapNewEnd = sentineled[a + 1][0];
    const gapOldStart = sentineled[a][1] + 1;
    const gapOldEnd = sentineled[a + 1][1];

    const pairCount = Math.min(gapNewEnd - gapNewStart, gapOldEnd - gapOldStart);
    for (let p = 0; p < pairCount; p++) {
      mapping[gapNewStart + p] = gapOldStart + p;
    }
  }

  return mapping;
}

/**
 * セル内改行（LF）ベースの差分ハイライトを適用する。
 * LCSで行を対応付けた上で、変更行は文字単位、追加行は全体をハイライトする。
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

  const hasChange = cleanNew !== cleanOld;
  let currentPos = 0;
  const lineMapping = buildOldLineMapping(oldLines, newLines);

  newLines.forEach((newLine, lineIdx) => {
    const oldLineIdx = lineMapping[lineIdx];

    if (oldLineIdx === -1) {
      // 新規追加行 — 全文字をハイライト
      if (newLine.length > 0) {
        richTextBuilder.setTextStyle(currentPos, currentPos + newLine.length, diffTextStyle);
      }
    } else {
      const oldLine = oldLines[oldLineIdx];
      if (newLine !== oldLine) {
        // 変更行 — 文字単位で差分検出
        for (let i = 0; i < newLine.length; i++) {
          if (i >= oldLine.length || newLine[i] !== oldLine[i]) {
            richTextBuilder.setTextStyle(currentPos + i, currentPos + i + 1, diffTextStyle);
          }
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
