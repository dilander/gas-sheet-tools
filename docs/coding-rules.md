# コーディング規約・プロジェクト仕様

## 技術スタック

- **ランタイム**: Google Apps Script (V8)
- **言語**: JavaScript（TypeScriptではない）
- **デプロイ**: clasp (`@google/clasp`) で push/pull
- **型情報**: JSDocで記述（`@param`, `@returns`, `@throws` 等）

## ファイル構成と責務

| ファイル | 責務 |
|---|---|
| `Config.js` | 共通定数（`CONFIG`オブジェクト）・ヘルパー関数。他ファイルから依存される |
| `Export.js` | Markdownエクスポート機能 |
| `Handlers.js` | セル編集時の差分ハイライト |
| `Menu.js` | カスタムメニュー登録・バックアップキャッシュ準備 |
| `NightlyTasks.js` | 定時バッチ（エクスポート＋バックアップ） |

## コーディングルール

### 基本

- `const` 優先。`let` は再代入が必要な場合のみ。`var` は使用禁止
- マジック値は `Config.js` の `CONFIG` オブジェクトに集約する（ハードコード禁止）
- フォルダ操作は `getParentFolder()` と `getOrCreateSubFolder()` を再利用する

### JSDoc

- 全関数にJSDocコメントを付与する（`@param`, `@returns`, `@throws`）
- ファイル先頭には `@fileoverview` を記述する

### GAS固有の制約

- `import` / `export` は使用不可。全関数はグローバルスコープ
- Node.js APIやブラウザAPIは使用不可。GAS固有のAPI（`SpreadsheetApp`, `DriveApp`, `CacheService` 等）を使う
- ファイル間の依存は実行順序で解決される。`Config.js` が最初に読み込まれる前提
- `SpreadsheetApp.getUi()` はバックグラウンド実行（トリガー）時にUIを返せない。try-catchで囲む
- `CacheService.getScriptCache()` の上限は100KB/キー、TTLは最大6時間（21600秒）

## デプロイ手順

```bash
clasp push    # ローカル → GAS
clasp pull    # GAS → ローカル
```
