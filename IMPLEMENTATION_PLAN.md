# 実装方針

## 1. システムアーキテクチャ

```
┌─────────────────────────────────┐
│    Google Sheets                 │
│  ┌──────────────────────────┐   │
│  │  検索条件シート           │   │
│  │ (開始日,終了日,キーワード) │   │
│  └──────────────────────────┘   │
│            ↓                     │
│  ┌──────────────────────────┐   │
│  │  結果シート               │   │
│  │ (メール情報,処理結果)     │   │
│  └──────────────────────────┘   │
└──────────────┬────────────────────┘
               │
        ┌──────┴──────┐
        ↓             ↓
    ┌─────────┐  ┌──────────┐
    │ Gmail   │  │ Google   │
    │ API     │  │ Drive    │
    │         │  │ API      │
    └─────────┘  └──────────┘
```

---

## 2. ファイル構成

```
copy-gmail-attachments-to-drive/
├── README.md                      # プロジェクト概要
├── REQUIREMENTS.md                # 要件定義（本ドキュメント）
├── IMPLEMENTATION_PLAN.md         # 実装方針（本ドキュメント）
├── src/
│   ├── main.gs                    # メイン処理・ UI実装
│   ├── gmail.gs                   # Gmail API ラッパー
│   ├── drive.gs                   # Google Drive API ラッパー
│   ├── sheet.gs                   # Google Sheets ラッパー
│   └── utils.gs                   # ユーティリティ関数
├── appsscript.json                # GAS マニフェストファイル
└── CHANGELOG.md                   # 更新履歴
```

---

## 3. 主要コンポーネント設計

### 3.1 main.gs

**責務**：

- UI制御（メニュー追加）
- 検索フロー管理
- 実行フロー管理
- エラーハンドリング

**主要関数**：

```javascript
function onOpen()
// Google Sheets 起動時にカスタムメニューを追加

function searchAndDisplay()
// 「検索して一覧表示」ボタン実行
// 1. 検索条件シートから入力値を取得
// 2. 入力値を検証
// 3. メール検索を実行
// 4. 結果シートにデータを書き込み

function executeAndCopy()
// 「選択対象をGoogleドライブにコピー」ボタン実行
// 1. 結果シートから保存対象=ONのRowを抽出
// 2. 各メールのファイルをコピー
// 3. 処理結果を結果シートに書き込み
```

---

### 3.2 gmail.gs

**責務**：Gmail API ラッパー

**主要関数**：

```javascript
function searchEmails(startDate, endDate, keyword)
// 入力:
//   - startDate (Date): 開始日
//   - endDate (Date): 終了日
//   - keyword (String): キーワード（複数の場合はカンマ区切り）
// 出力: メールスレッドの配列

function extractAttachments(thread)
// 入力: メールスレッド
// 出力:
// {
//   subject: "メールタイトル",
//   date: Date,
//   attachments: [{name, blobSource}]
// }

function getPdfAttachments(thread)
// 入力: メールスレッド
// 出力: PDF ファイルのみを抽出した配列
```

---

### 3.3 drive.gs

**責務**：Google Drive API ラッパー

**主要関数**：

```javascript
function getFolderByPath(folderPath)
// 入力: フォルダパス （例："/プロジェクト/PDFフォルダ"）
// 出力: フォルダID または 404 エラー

function createFolderByPath(folderPath)
// 入力: フォルダパス
// 出力: 作成または既存フォルダの ID
// （存在しない親フォルダも再帰的に作成）

function copyFileToFolder(file, folderId, newFileName)
// 入力:
//   - file: Blob (ファイルデータ)
//   - folderId: 保存先フォルダID
//   - newFileName: 保存ファイル名
// 出力:
// {
//   success: boolean,
//   fileId: String,
//   error: String (失敗時のみ)
// }

function fileExistsInFolder(folderId, fileName)
// 入力: フォルダID、ファイル名
// 出力: true/false
```

---

### 3.4 sheet.gs

**責務**：Google Sheets API ラッパー

**主要関数**：

```javascript
function getSearchConditions()
// 出力:
// {
//   startDate: Date,
//   endDate: Date,
//   keywords: [String],
//   folderPath: String
// }

function clearResultSheet()
// 結果シートをクリア（ヘッダーは保持）

function writeResults(resultsArray)
// 入力: 結果オブジェクト配列
// {
//   targetCheckbox: boolean,
//   title: String,
//   date: Date,
//   attachmentNames: String (カンマ区切り),
//   saveFileName: String,
//   saveFolder: String,
//   result: String (OK/SKIP/エラー),
//   fileLink: String
// }

function addResultRow(rowData)
// 結果シートに1行追加
```

---

### 3.5 utils.gs

**責務**：ユーティリティ関数

**主要関数**：

```javascript
function validateDateInput(dateString)
// 入力: 日付文字列 (YYYY-MM-DD)
// 出力: Date または throw Error

function formatDateTime(date)
// 入力: Date
// 出力: YYYY-MM-DD HH:MM:SS 形式の文字列

function ensurePdfExtension(fileName)
// 入力: ファイル名
// 出力: .pdf 拡張子が付いたファイル名

function parseKeywords(keywordString)
// 入力: キーワード文字列（カンマ区切り）
// 出力: キーワード配列

function createDriveLink(fileId)
// 入力: Google Drive ファイルID
// 出力: https://drive.google.com/file/d/{fileId}/view
```

---

## 4. 処理フロー詳細

### 4.1 検索フロー（searchAndDisplay）

```
START
  ↓
検索条件シートから入力値を取得
  ↓
入力値の検証（日付形式、必須項目）
  ↓
[エラー？] → YES → エラーダイアログ表示 → END
  ↓ NO
Gmail検索クエリを構築
  ↓
gmail.searchEmails() を実行
  ↓
各メールについて:
  - タイトル、日時を取得
  - PDF添付ファイルを抽出
  ↓
結果シートをクリア
  ↓
抽出したメール情報を結果シートに書き込み
  ↓
完了ダイアログ表示（件数）
  ↓
END
```

### 4.2 実行フロー（executeAndCopy）

```
START
  ↓
結果シートから保存対象=ON の Rowを抽出
  ↓
各Rowについて:
  ┌─────────────────────────────────────┐
  │ メールID、ファイル名を取得          │
  │           ↓                         │
  │ 保存先フォルダパスを取得            │
  │           ↓                         │
  │ フォルダIDに変換（未作成なら作成）   │
  │           ↓                         │
  │ [ファイル名が既に存在？]             │
  │    ↓YES → "SKIP" と表示             │
  │    ↓NO                              │
  │ ファイルをコピー                    │
  │           ↓                         │
  │ [コピー成功？]                      │
  │    ↓YES → "OK" と表示、リンク生成   │
  │    ↓NO  → "エラー" と表示          │
  │           ↓                         │
  │ 処理結果を結果シートに書き込み      │
  └─────────────────────────────────────┘
  ↓
完了ダイアログ表示（成功数、スキップ数、エラー数）
  ↓
END
```

---

## 5. Gmail 検索クエリの構築

GmailAPI の `search()` メソッドで以下のクエリを使用：

```
基本クエリ: is:attachment after:YYYY-MM-DD before:YYYY-MM-DD+1

キーワード追加の場合:
  - キーワード1: subject:(keyword1)
  - キーワード2: subject:(keyword1 OR keyword2)

例：
  is:attachment after:2026-01-01 before:2026-02-01 subject:(請求書 OR 報告書)
```

---

## 6. フォルダパス解析アルゴリズム

### パスの形式

```
絶対パス：/フォルダA/フォルダB/フォルダC
ルート：/ または 空欄
```

### 解析ロジック

```javascript
function resolveFolderPath(pathString) {
  if (!pathString || pathString === "/") {
    return DriveApp.getRootFolder();
  }

  const parts = pathString.split("/").filter((p) => p.length > 0);
  let currentFolder = DriveApp.getRootFolder();

  for (const folderName of parts) {
    try {
      const subfolders = currentFolder.getFoldersByName(folderName);
      if (subfolders.hasNext()) {
        currentFolder = subfolders.next();
      } else {
        // フォルダが見つからない → 作成
        currentFolder = currentFolder.createFolder(folderName);
      }
    } catch (error) {
      throw new Error(`フォルダ作成失敗: ${folderName}`);
    }
  }

  return currentFolder;
}
```

---

## 7. ファイル名重複チェック

Google Drive APIでは、同じフォルダ内での同名ファイルを許可する仕様のため、チェック処理を実装：

```javascript
function fileExistsInFolder(folderId, fileName) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByName(fileName);
  return files.hasNext();
}
```

---

## 8. エラーハンドリング戦略

### Try-Catch ブロックの配置

| レイヤー | 対象                  | 動作                            |
| -------- | --------------------- | ------------------------------- |
| main.gs  | API全体               | エラーダイアログ表示、処理中止  |
| gmail.gs | メール検索            | スタックトレース記録、Null 返却 |
| drive.gs | フォルダ/ファイル操作 | エラーメッセージ返却（行継続）  |
| sheet.gs | シート読み書き        | エラーログ出力                  |

### ユーザーへの通知

- **入力エラー**：入力値チェックで即座に通知
- **API エラー**：スクリプト実行中にコンソール出力
- **処理結果**：結果シートに「OK」「SKIP」「エラー」と記載

---

## 9. 開発段階

### フェーズ1：基盤構築（Week 1）

- [ ] appsscript.json の設定
- [ ] utils.gs：ユーティリティ関数実装
- [ ] sheet.gs：Sheets ラッパー実装

### フェーズ2：API ラッパー実装（Week 1-2）

- [ ] gmail.gs：Gmail API ラッパー実装
- [ ] drive.gs：Drive API ラッパー実装

### フェーズ3：UI・メイン処理（Week 2）

- [ ] main.gs：UI・フロー実装
- [ ] 各関数のユニットテスト（手動）

### フェーズ4：統合テスト・ドキュメント（Week 3）

- [ ] 統合テスト
- [ ] README.md の詳細記述
- [ ] CHANGELOG.md の記述

---

## 10. テスト計画

### 手動テストケース

| テスト項目       | 入力例                      | 期待結果                     |
| ---------------- | --------------------------- | ---------------------------- |
| 期間検索         | 2026-01-01 ～ 2026-01-31    | 該当メール一覧表示           |
| キーワード検索   | "請求書"                    | タイトルに含むメールのみ表示 |
| 無効な日付       | 2026-13-01                  | エラーダイアログ表示         |
| フォルダ自動作成 | `/新規フォルダ/PDFフォルダ` | フォルダ自動作成             |
| ファイル名重複   | 既存ファイル名              | SKIP と表示                  |
| ファイルコピー   | PDF ファイル                | OK と表示、リンク生成        |

---

## 11. 今後の拡張予定（Out-of-scope）

- [ ] 添付ファイルタイプのフィルタ（PDF のみから複数形式対応）
- [ ] バッチ処理の自動スケジューリング（Google Cloud Scheduler）
- [ ] 処理ログの Google Sheets での永続化
- [ ] UI の高度な検索フィルタ（複数条件 AND/OR）

---
