# Copilot 説明: Gmail 添付ファイルを Drive にコピー

## プロジェクト概要

指定された日付範囲内の Gmail から PDF 添付ファイルを抽出し、Google Drive フォルダにコピーする Google Apps Script ツール。アプリは Google Sheets を使用したスプレッドシートベースの UI を提供し、2つのメインシートがあります:

- **検索条件**: 日付範囲、キーワード、保存先フォルダパスを入力するシート
- **結果**: 検索結果をチェックボックス付きで表示するシート

## アーキテクチャとデータフロー

### コアモジュール構造（TypeScript）

**`src/main.ts`** - UI オーケストレーションレイヤー

- `onOpen()`: カスタムメニュー「Gmail添付ファイル」とアクションアイテムを作成
- `searchAndDisplay()`: メール検索と結果シートの入力ワークフロー
- `executeAndCopy()`: チェック済み行の処理とファイルコピーワークフロー
- `initializeSheets()`: デフォルトヘッダー/データで必要なシートを作成

**`src/gmail.ts`** - Gmail API ラッパー

- `searchEmails()`: Gmail クエリ構文（`is:attachment after:... before:...`）を使用してスレッドを検索
- `buildSearchQuery()`: 日付範囲とキーワードフィルタで Gmail 検索クエリを構築
- `extractEmailAndAttachments()`: メールスレッドからサブジェクト、日付、PDF を抽出
- `getPdfAttachments()`: メール添付ファイルから PDF ファイルをフィルタリング
- `SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()` によるタイムゾーン対応の日付処理

**`src/drive.ts`** - Google Drive API ラッパー

- `resolveFolderPath()`: パス文字列（例：`/year/2026/Jan`）からフォルダ階層をナビゲート/作成
- `fileExistsInFolder()`: 重複ファイルの上書きを防止
- `copyFileToFolder()`: エラーハンドリング付きで Drive にブロブをコピー
- `ensurePdfExtension()`: ファイル名に `.pdf` 拡張子を確保

**`src/sheet.ts`** - Google Sheets API ラッパー

- `getSearchSheet()` / `getResultsSheet()`: 遅延初期化パターン
- `getSearchConditions()`: 検索シートの 2～5 行を型付きオブジェクトに読み込む
- `getTargetRowsData()`: 結果シートのチェック済み行を読み込む
- `addResultRows()`: Drive リンク用の数式を含む結果データをバッチ挿入
- チェックボックス状態と共有 Drive リンク用の数式生成を処理

**`src/utils.ts`** - 共通ユーティリティ

- `validateDateInput()`: YYYY-MM-DD 文字列または Date オブジェクトを解析
- `formatDateTime()`: 表示/ログ用に日付をフォーマット
- `logDebug()` / `logError()`: コンテキストオブジェクト付きログ

## 主要開発パターン

### ビルドとデプロイ

```bash
# 開発ワークフロー
yarn build          # TypeScript → JavaScript (clasp push は JS が必要)
yarn watch          # 開発用ウォッチモード

# Google Apps Script にデプロイ
yarn push             # コンパイルして GAS プロジェクトにコードをプッシュ
yarn logs          # clasp 経由で実行ログを表示
```

**重要**: TypeScript ソースファイルを編集した後は、必ず `yarn push` を実行して Google Apps Script にコンパイルとデプロイを行います。このコマンドを実行するまで、GAS エディタに変更が反映されません。

`appsscript.json` は必要な OAuth スコープを宣言しています:

- `gmail.readonly`（メール検索、添付ファイル読み込み）
- `spreadsheets`（シート読み書き）
- `drive`（フォルダ作成、ファイルコピー）

### エラーハンドリングパターン

すべての非同期 Google Apps Script API 呼び出しは try-catch でラップされます:

```typescript
try {
  // API 呼び出し
  logDebug("ステップコンテキスト", { detailKey: value });
  return result;
} catch (error) {
  const err = error as Error;
  logError("操作名", err.message);
  throw new Error(`ユーザーフレンドリーなメッセージ: ${err.message}`);
}
```

エラーは `main.ts` の関数にバブルアップされ、`SpreadsheetApp.getUi().alert()` 経由でアラートを表示します。

### 型定義

モジュール間のデータコントラクト用のカスタム型:

- `EmailData`: {subject, date, attachments[], messageId}
- `SearchConditions`: {startDate, endDate, keywords[], folderPath}
- `ResultRow`: 結果シートの列に対応
- `CopyResult`: {success, fileId?, error?}

### タイムゾーン処理

**重要**: すべての日付操作はスプレッドシートのタイムゾーンを考慮します:

- シートから入力される日付はスプレッドシートのタイムゾーンに基づくローカル時間と見なされます
- Gmail クエリ日付は `Utilities.formatDate(date, timezone, "yyyy/MM/dd")` でフォーマットされます
- タイムゾーン変換なし—日付はシステム全体を通じて直接流れます

## シート UI 構造

### 検索シート（検索条件）

```
| 項目              | 値         |
|-------------------|----------|
| 期間（開始日）    | 2026-01-01 |
| 期間（終了日）    | 2026-01-31 |
| フィルタ文字列    | (optional keywords, comma-separated) |
| 保存先フォルダパス | /project/pdfs |
```

### 結果シート（結果）

列: 保存対象 (チェックボックス) | メールタイトル | 受信日時 | 添付ファイル名 | 保存ファイル名 | 保存先フォルダ名 | 処理結果 | ファイルリンク

## 統合ポイント & 依存関係

- **GmailApp**: スレッド検索、メールメタデータ、添付ファイル Blob
- **DriveApp**: フォルダナビゲーション、ファイル作成、重複検出
- **SpreadsheetApp**: シートアクセス、範囲操作、タイムゾーン情報
- **Utilities**: 日付フォーマット、型チェック

## 重要な考慮事項

1. **PDF のみフィルタリング**: PDF 以外の添付ファイルは静かにスキップされます
2. **タイムゾーン認識**: シートの日付文字列はスプレッドシートのタイムゾーンと見なされます
3. **重複防止**: 同じフォルダ内に同じ名前のファイルがある場合はスキップされます（上書きされません）
4. **バッチ操作**: `addResultRows()` はパフォーマンスのためバッチ挿入を使用します
5. **検索の制限**: Gmail 検索構文は日付範囲をサポートしていますが、細かい時刻フィルタリングはできません
