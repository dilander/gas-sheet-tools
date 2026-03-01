/**
 * @fileoverview Markdownエクスポート機能
 * スプレッドシートの全シートをMarkdown形式に変換し、
 * Google Drive上のexportフォルダにテキストファイルとして出力する。
 *
 * 列名による特殊処理:
 * - "Markdown" を含む列: 見出しレベルをH4基準にオフセット調整し、インデント出力
 * - "Python" / "Handlebars" / "Mermaid" を含む列: コードブロックで囲んで出力
 * - "無視" を含む列: エクスポート対象から除外
 */

/**
 * メインのエクスポート処理。全シートをMarkdownに変換してDriveに保存する。
 * メニューおよびナイトリービルドから呼び出される。
 */
function exportSheetsToMarkdown() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ssName = spreadsheet.getName();
  const ssId = spreadsheet.getId();
  const timestamp = getTimestamp();

  const markdown = buildMarkdown(spreadsheet, ssName, timestamp);

  const parentFolder = getParentFolder(ssId);
  const exportFolder = getOrCreateSubFolder(parentFolder, CONFIG.FOLDER.EXPORT);
  const fileName = `${ssName}_export_${timestamp}${CONFIG.EXPORT.EXTENSION}`;

  try {
    const file = exportFolder.createFile(fileName, markdown, MimeType.PLAIN_TEXT);
    console.log(`[exportSheetsToMarkdown]: エクスポート完了: ${fileName}`);
    showDownloadLink(file);
  } catch (e) {
    console.error(`[exportSheetsToMarkdown]: エクスポートに失敗しました: ${e.toString()}`);
    try {
      SpreadsheetApp.getUi().alert("エラー: " + e.toString());
    } catch (_) {}
  }
}

/**
 * スプレッドシート全体をMarkdown文字列に変換する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - 対象スプレッドシート
 * @param {string} ssName - スプレッドシート名（ヘッダー表示用）
 * @param {string} timestamp - エクスポート日時の文字列
 * @returns {string} Markdown形式のテキスト
 */
function buildMarkdown(spreadsheet, ssName, timestamp) {
  let md = `# Source: ${ssName}\n`;
  md += `> Export Date: ${timestamp}\n\n`;
  md += `---\n\n`;

  spreadsheet.getSheets().forEach(sheet => {
    md += buildSheetMarkdown(sheet);
  });

  return md;
}

/**
 * 単一シートをMarkdown文字列に変換する。
 * 1列目をタイトル（H3）として扱い、2列目以降をリスト項目として出力する。
 * タイトル列が空の行は直前のタイトルを引き継ぐ。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @returns {string} Markdown形式のテキスト（データなしの場合は空文字列）
 */
function buildSheetMarkdown(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return "";

  const headers = data[0];
  const validCols = headers.reduce((acc, h, j) => {
    if (h.toString().indexOf(CONFIG.IGNORE_KEYWORD) === -1) acc.push(j);
    return acc;
  }, []);

  if (validCols.length === 0) return "";

  let md = `## Sheet: ${sheet.getName()}\n\n`;
  let lastTitle = "";

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let title = row[validCols[0]];
    if (!title) title = lastTitle;
    else lastTitle = title;

    md += `### ${title || "(Untitled)"}\n`;

    for (let k = 1; k < validCols.length; k++) {
      const colIdx = validCols[k];
      const headerName = headers[colIdx].toString();
      const cellValue = row[colIdx];
      if (cellValue === "") continue;

      md += formatCellValue(headerName, cellValue);
    }
    md += "\n";
  }

  return md;
}

/**
 * セル値を列名に応じたMarkdown書式に変換する。
 * @param {string} headerName - 列名
 * @param {*} cellValue - セルの値
 * @returns {string} フォーマット済みのMarkdown行
 */
function formatCellValue(headerName, cellValue) {
  if (headerName.indexOf("Markdown") !== -1) {
    return formatMarkdownCell(headerName, cellValue);
  }

  const codeMatch = headerName.match(CONFIG.CODE_LANGUAGES);
  if (codeMatch) {
    return formatCodeCell(headerName, cellValue, codeMatch[1].toLowerCase());
  }

  const cleanHeaderName = removeMetadata(headerName);
  const strValue = cellValue.toString();
  if (strValue.indexOf('\n') !== -1) {
    const indented = strValue.split('\n').map(line => '  ' + line).join('\n');
    return `- ${cleanHeaderName}: \n${indented}\n`;
  }
  return `- ${cleanHeaderName}: ${strValue}\n`;
}

/**
 * 列名から処理に使用するメタデータのみを削除する。
 * 削除対象: "（無視）"、"（Markdown）"、"（Python）"、"（Handlebars）"、"（Mermaid）"、"（YAML）"
 * その他の括弧付きメタデータ（例: "（調査で判明）"、"（18禁解禁時）"）は保持する。
 * @param {string} headerName - 元の列名
 * @returns {string} 処理用メタデータを除去した列名
 */
function removeMetadata(headerName) {
  return headerName
    .replace(/[（(]無視[）)]/g, '')
    .replace(/[（(]Markdown[）)]/g, '')
    .replace(/[（(]Python[）)]/g, '')
    .replace(/[（(]Handlebars[）)]/g, '')
    .replace(/[（(]Mermaid[）)]/g, '')
    .replace(/[（(]YAML[）)]/g, '')
    .trim();
}

/**
 * Markdown列の値を処理する。見出しレベルをH4基準にオフセット調整し、
 * 2スペースインデントで出力する。
 * @param {string} headerName - 列名
 * @param {*} cellValue - セルの値（Markdownテキスト）
 * @returns {string} オフセット・インデント適用済みのMarkdown行
 */
function formatMarkdownCell(headerName, cellValue) {
  let lines = cellValue.toString().split('\n');
  let minHashes = Infinity;

  for (const line of lines) {
    const match = line.match(/^(#+)\s/);
    if (match) minHashes = Math.min(minHashes, match[1].length);
  }

  if (minHashes !== Infinity) {
    const offset = CONFIG.HEADING_BASE_LEVEL - minHashes;
    lines = lines.map(line => {
      const match = line.match(/^(#+)(\s.*)/);
      if (match) {
        const newCount = Math.max(1, match[1].length + offset);
        return '#'.repeat(newCount) + match[2];
      }
      return line;
    });
  }

  const cleanHeaderName = removeMetadata(headerName);
  const indented = lines.map(line => '  ' + line).join('\n');
  return `- ${cleanHeaderName}: \n${indented}\n`;
}

/**
 * コード列の値をコードブロックで囲んで出力する。
 * @param {string} headerName - 列名
 * @param {*} cellValue - セルの値（コードテキスト）
 * @param {string} lang - コードブロックの言語指定（小文字）
 * @returns {string} コードブロック付きのMarkdown行
 */
function formatCodeCell(headerName, cellValue, lang) {
  const cleanHeaderName = removeMetadata(headerName);
  const codeBlock = `\`\`\`${lang}\n${cellValue}\n\`\`\``;
  const indented = codeBlock.split('\n').map(line => '  ' + line).join('\n');
  return `- ${cleanHeaderName}: \n${indented}\n`;
}

// ============================================================
// Separate モード
// ============================================================

/**
 * Separateモードのエクスポート処理。3ファイルに分割して出力する。
 * - *_separate.md: YAML・Handlebars・Pythonシグネチャ・その他通常列（Mermaid除外）
 * - *_functions.py: Python全関数（コメント区切り）
 * - *_floormaps.py: Mermaidテキストを辞書型変数として格納
 */
function exportSeparateMode() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ssName = spreadsheet.getName();
  const ssId = spreadsheet.getId();
  const timestamp = getTimestamp();

  const parentFolder = getParentFolder(ssId);
  const exportFolder = getOrCreateSubFolder(parentFolder, CONFIG.FOLDER.EXPORT);

  const separateMd = buildSeparateMarkdown(spreadsheet, ssName, timestamp);
  const functionsPy = buildFunctionsPy(spreadsheet);
  const floormapsPy = buildFloormapsPy(spreadsheet);

  const blobs = [
    Utilities.newBlob(separateMd, MimeType.PLAIN_TEXT, `${ssName}_export_${timestamp}_separate.md`),
    Utilities.newBlob(functionsPy, MimeType.PLAIN_TEXT, `${ssName}_export_${timestamp}_functions.py`),
    Utilities.newBlob(floormapsPy, MimeType.PLAIN_TEXT, `${ssName}_export_${timestamp}_floormaps.py`),
  ];

  try {
    const zipBlob = Utilities.zip(blobs, `${ssName}_export_${timestamp}_separate.zip`);
    const zipFile = exportFolder.createFile(zipBlob);
    console.log(`[exportSeparateMode]: エクスポート完了: ${zipFile.getName()}`);
    showDownloadLink(zipFile);
  } catch (e) {
    console.error(`[exportSeparateMode]: エクスポートに失敗しました: ${e.toString()}`);
    try {
      SpreadsheetApp.getUi().alert("エラー: " + e.toString());
    } catch (_) {}
  }
}

/**
 * Separateモード用Markdownを構築する。
 * YAML・Handlebars列はコードブロック、Python列はdef行のみ、Mermaid列は除外。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - 対象スプレッドシート
 * @param {string} ssName - スプレッドシート名
 * @param {string} timestamp - エクスポート日時
 * @returns {string} Markdown文字列
 */
function buildSeparateMarkdown(spreadsheet, ssName, timestamp) {
  let md = `# Source: ${ssName}\n`;
  md += `> Export Date: ${timestamp}\n`;
  md += `> Mode: Separate (Claude / ChatGPT)\n\n`;
  md += `---\n\n`;

  spreadsheet.getSheets().forEach(sheet => {
    md += buildSheetSeparateMarkdown(sheet);
  });

  return md;
}

/**
 * 単一シートをSeparateモード用Markdownに変換する。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @returns {string} Markdown文字列
 */
function buildSheetSeparateMarkdown(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return "";

  const headers = data[0];
  const validCols = headers.reduce((acc, h, j) => {
    const name = h.toString();
    if (name.indexOf(CONFIG.IGNORE_KEYWORD) === -1 && !(/Mermaid/i.test(name))) {
      acc.push(j);
    }
    return acc;
  }, []);

  if (validCols.length === 0) return "";

  let md = `## Sheet: ${sheet.getName()}\n\n`;
  let lastTitle = "";

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let title = row[validCols[0]];
    if (!title) title = lastTitle;
    else lastTitle = title;

    md += `### ${title || "(Untitled)"}\n`;

    for (let k = 1; k < validCols.length; k++) {
      const colIdx = validCols[k];
      const headerName = headers[colIdx].toString();
      const cellValue = row[colIdx];
      if (cellValue === "") continue;

      md += formatSeparateCellValue(headerName, cellValue);
    }
    md += "\n";
  }

  return md;
}

/**
 * Separateモード用のセル値フォーマット。
 * Python列はdef行のみ抽出、それ以外はCombineモードと同じ処理。
 * @param {string} headerName - 列名
 * @param {*} cellValue - セルの値
 * @returns {string} フォーマット済みのMarkdown行
 */
function formatSeparateCellValue(headerName, cellValue) {
  if (/Python/i.test(headerName) && headerName.indexOf("実行サンプル") === -1) {
    return formatPythonSignature(headerName, cellValue);
  }

  return formatCellValue(headerName, cellValue);
}

/**
 * Python列からdef行（関数シグネチャ）とdocstringを抽出してPythonコードブロックで出力する。
 * @param {string} headerName - 列名
 * @param {*} cellValue - セルの値（Pythonコード）
 * @returns {string} シグネチャ+docstringのMarkdownコードブロック
 */
function formatPythonSignature(headerName, cellValue) {
  const signatures = extractPythonSignatures(cellValue.toString());
  if (signatures.length === 0) return "";

  const cleanHeaderName = "関数シグネチャ";
  const codeContent = signatures.map(sig => {
    if (sig.docstringBlock) {
      return `${sig.def}\n${sig.docstringBlock}`;
    }
    return sig.def;
  }).join('\n\n');
  const codeBlock = `\`\`\`python\n${codeContent}\n\`\`\``;
  const indented = codeBlock.split('\n').map(line => '  ' + line).join('\n');
  return `- ${cleanHeaderName}:\n${indented}\n`;
}

/**
 * Pythonコードからトップレベルのdef文とその直後のdocstringブロックを抽出する。
 * 複数行にまたがるdef文（引数が長い場合）にも対応。
 * docstringは元のインデント・改行を保持した状態で返す。
 * @param {string} code - Pythonコード文字列
 * @returns {Array<{def: string, docstringBlock: string}>} シグネチャとdocstringブロックの配列
 */
function extractPythonSignatures(code) {
  const lines = code.split('\n');
  const results = [];

  for (let i = 0; i < lines.length; i++) {
    if (!/^def\s+/.test(lines[i])) continue;

    // def文を収集（複数行の場合は ): or ) -> ...: まで連結）
    const defParts = [lines[i].trim()];
    let endIdx = i;
    if (lines[i].trim().slice(-1) !== ':') {
      for (let k = i + 1; k < lines.length; k++) {
        defParts.push(lines[k].trim());
        endIdx = k;
        if (lines[k].trim().slice(-1) === ':') break;
      }
    }
    const defLine = defParts.join(' ');

    // docstringブロックを探す（元のインデント・改行を保持）
    let docstringBlock = "";
    const nextIdx = endIdx + 1;
    if (nextIdx < lines.length) {
      const trimmed = lines[nextIdx].trim();
      const quoteMatch = trimmed.match(/^("""|''')/);
      if (quoteMatch) {
        const quote = quoteMatch[1];
        if (trimmed.indexOf(quote, quote.length) !== -1) {
          // 1行docstring: """text""" → そのまま保持
          docstringBlock = lines[nextIdx];
        } else {
          // 複数行docstring: 開始行から閉じクォートの行まで保持
          const docLines = [lines[nextIdx]];
          for (let j = nextIdx + 1; j < lines.length; j++) {
            docLines.push(lines[j]);
            if (lines[j].trim().indexOf(quote) !== -1) break;
          }
          docstringBlock = docLines.join('\n');
        }
      }
    }

    results.push({ def: defLine, docstringBlock: docstringBlock });
    i = endIdx;
  }

  return results;
}

/**
 * 全シートからPython列のコードを収集し、docstringを除去してコメント区切りで連結する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - 対象スプレッドシート
 * @returns {string} Python コード文字列
 */
function buildFunctionsPy(spreadsheet) {
  const blocks = [];

  spreadsheet.getSheets().forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    const headers = data[0];
    const pythonCols = headers.reduce((acc, h, j) => {
      const name = h.toString();
      if (/Python/i.test(name) && name.indexOf(CONFIG.IGNORE_KEYWORD) === -1 && name.indexOf("実行サンプル") === -1) {
        acc.push(j);
      }
      return acc;
    }, []);

    if (pythonCols.length === 0) return;

    let lastTitle = "";
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      let title = row[0];
      if (!title) title = lastTitle;
      else lastTitle = title;

      for (const colIdx of pythonCols) {
        const cellValue = row[colIdx];
        if (cellValue === "") continue;
        const stripped = stripDocstrings(cellValue.toString());
        blocks.push(`# === ${title || "(Untitled)"} ===\n${stripped}`);
      }
    }
  });

  return blocks.join('\n\n\n');
}

/**
 * Pythonコードからdocstringを除去する。
 * def文直後の"""..."""または'''...'''ブロックを削除する。
 * 複数行にまたがるdef文にも対応。
 * @param {string} code - Pythonコード文字列
 * @returns {string} docstring除去済みのコード
 */
function stripDocstrings(code) {
  const lines = code.split('\n');
  const result = [];
  let i = 0;

  while (i < lines.length) {
    result.push(lines[i]);

    // def行を検出（トップレベル・ネスト両方対象）
    if (/^\s*def\s+/.test(lines[i])) {
      // 複数行defの場合、末尾が : になるまで出力を続ける
      while (lines[i].trim().slice(-1) !== ':' && i + 1 < lines.length) {
        i++;
        result.push(lines[i]);
      }

      // def文終了後の次の行でdocstringを探す
      const nextIdx = i + 1;
      if (nextIdx < lines.length) {
        const trimmed = lines[nextIdx].trim();
        const quoteMatch = trimmed.match(/^("""|''')/);
        if (quoteMatch) {
          const quote = quoteMatch[1];
          if (trimmed.indexOf(quote, quote.length) !== -1) {
            i = nextIdx + 1;
            continue;
          }
          let j = nextIdx + 1;
          while (j < lines.length) {
            if (lines[j].trim().indexOf(quote) !== -1) {
              i = j + 1;
              break;
            }
            j++;
          }
          if (j >= lines.length) i = j;
          continue;
        }
      }
    }

    i++;
  }

  return result.join('\n');
}

/**
 * 全シートからMermaid列のテキストを収集し、辞書型変数として格納する。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - 対象スプレッドシート
 * @returns {string} Python コード文字列（FLOORMAPS辞書）
 */
function buildFloormapsPy(spreadsheet) {
  const entries = [];

  spreadsheet.getSheets().forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    const headers = data[0];
    const mermaidCols = headers.reduce((acc, h, j) => {
      if (/Mermaid/i.test(h.toString()) && h.toString().indexOf(CONFIG.IGNORE_KEYWORD) === -1) {
        acc.push(j);
      }
      return acc;
    }, []);

    if (mermaidCols.length === 0) return;

    let lastTitle = "";
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      let title = row[0];
      if (!title) title = lastTitle;
      else lastTitle = title;

      for (const colIdx of mermaidCols) {
        const cellValue = row[colIdx];
        if (cellValue === "") continue;
        const key = (title || "(Untitled)").toString().replace(/"/g, '\\"');
        entries.push(`    "${key}": """${cellValue}"""`);
      }
    }
  });

  if (entries.length === 0) return "FLOORMAPS = {}\n";
  return `FLOORMAPS = {\n${entries.join(',\n')},\n}\n`;
}

/**
 * エクスポート完了後にダウンロードリンクをモーダルダイアログで表示する。
 * バックグラウンド実行時（UIなし）はスキップする。
 * @param {GoogleAppsScript.Drive.File} file - エクスポート済みファイル
 */
function showDownloadLink(file) {
  try {
    const ui = SpreadsheetApp.getUi();
    if (!ui) return;

    const downloadUrl = `https://drive.google.com/uc?export=download&id=${file.getId()}`;
    const html = `
      <html>
        <body style="font-family: sans-serif; font-size: 14px; text-align: center; padding: 20px;">
          <p>エクスポート完了</p>
          <p style="font-size: 12px; color: gray;">保存先: ${file.getName()}</p>
          <br>
          <a href="${downloadUrl}" target="_blank"
             style="display: inline-block; padding: 10px 20px; background: #4285f4; color: white; text-decoration: none; border-radius: 4px; font-weight: bold;"
             onclick="setTimeout(function(){ google.script.host.close(); }, 1500);">
             ファイルを保存
          </a>
          <script>
            (function() {
              const link = document.createElement('a');
              link.href = "${downloadUrl}"; link.target = "_blank";
              document.body.appendChild(link);
              link.click();
            })();
          </script>
        </body>
      </html>
    `;
    const output = HtmlService.createHtmlOutput(html).setHeight(190).setWidth(350);
    ui.showModalDialog(output, "Export Complete");
  } catch (e) {
    // バックグラウンド実行時はUIが使えないためスキップ
  }
}
