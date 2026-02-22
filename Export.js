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
    showDownloadLink(file);
  } catch (e) {
    console.error("exportSheetsToMarkdown failed: " + e.toString());
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

  return `- **${headerName}**: ${cellValue}\n`;
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

  const indented = lines.map(line => '  ' + line).join('\n');
  return `- **${headerName}**: \n${indented}\n`;
}

/**
 * コード列の値をコードブロックで囲んで出力する。
 * @param {string} headerName - 列名
 * @param {*} cellValue - セルの値（コードテキスト）
 * @param {string} lang - コードブロックの言語指定（小文字）
 * @returns {string} コードブロック付きのMarkdown行
 */
function formatCodeCell(headerName, cellValue, lang) {
  const codeBlock = `\`\`\`${lang}\n${cellValue}\n\`\`\``;
  const indented = codeBlock.split('\n').map(line => '  ' + line).join('\n');
  return `- **${headerName}**: \n${indented}\n`;
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
              setTimeout(function(){ google.script.host.close(); }, 5000);
            })();
          </script>
        </body>
      </html>
    `;
    const output = HtmlService.createHtmlOutput(html).setHeight(190).setWidth(350);
    ui.showModalDialog(output, "Export Complete");
  } catch (e) {
    console.log("Background build detected. Skipping UI.");
  }
}
