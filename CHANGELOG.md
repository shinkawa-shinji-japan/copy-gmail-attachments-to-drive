# CHANGELOG

すべての注目すべき変更がこのプロジェクトに記録されます。

形式は[Keep a Changelog](https://keepachangelog.com/ja/)に基づいています。
そして、このプロジェクトは[セマンティックバージョニング](https://semver.org/ja/)に従います。

---

## [1.0.0] - 2026-01-17

### 追加

#### コア機能

- ✨ Gmail から指定期間内のメール検索機能
- ✨ メールに含まれる PDF 添付ファイルの自動抽出
- ✨ メールタイトルによるキーワードフィルタリング
- ✨ Google Drive への一括ファイルコピー機能
- ✨ Google Drive フォルダの自動作成機能
- ✨ ファイル名重複チェックとスキップ機能
- ✨ コピー後の Google Drive ファイルリンク自動生成

#### UI・UX

- 🎨 Google Sheets による見やすいUI
  - 検索条件入力シート
  - 検索結果・処理結果表示シート
- 🎨 カスタムメニュー（Google Sheets 統合）
  - 初期化（シート作成）
  - 検索して一覧表示
  - 選択対象をGoogleドライブにコピー
- 🎨 チェックボックスによる保存対象選択
- 🎨 ファイル名・保存先フォルダのカスタマイズ機能

#### API ラッパー関数

- 📚 Gmail API ラッパー（gmail.gs）
  - `searchEmails()` - メール検索
  - `extractEmailAndAttachments()` - メール情報抽出
  - `getPdfAttachments()` - PDF フォルダ抽出
- 📚 Google Drive API ラッパー（drive.gs）
  - `resolveFolderPath()` - フォルダパス解析・作成
  - `fileExistsInFolder()` - ファイル存在チェック
  - `copyFileToFolder()` - ファイルコピー
  - `copyMultipleFiles()` - 複数ファイルコピー
- 📚 Google Sheets API ラッパー（sheet.gs）
  - `getSearchSheet()` / `getResultsSheet()` - シート取得・作成
  - `initializeSearchSheet()` / `initializeResultsSheet()` - シート初期化
  - `getSearchConditions()` - 検索条件取得
  - `addResultRow()` / `addResultRows()` - 結果書き込み
  - `getTargetRowsData()` - 処理対象データ取得

#### ユーティリティ

- 🛠 `validateDateInput()` - 日付入力検証
- 🛠 `formatDateTime()` - 日付フォーマット
- 🛠 `ensurePdfExtension()` - PDF 拡張子確認
- 🛠 `parseKeywords()` - キーワードパース
- 🛠 `createDriveLink()` - Google Drive リンク生成
- 🛠 `logDebug()` / `logWarn()` / `logError()` - ロギング関数

#### ドキュメント

- 📖 要件定義書（REQUIREMENTS.md）
- 📖 実装方針書（IMPLEMENTATION_PLAN.md）
- 📖 セットアップ・デプロイガイド（SETUP_GUIDE.md）
- 📖 README（プロジェクト概要）

#### 開発環境

- 🔧 appsscript.json マニフェスト設定
- 🔧 package.json clasp 統合設定
- 🔧 npm スクリプト
  - `npm run push` - コードをプッシュ
  - `npm run pull` - コードをプル
  - `npm run open` - エディタを開く
  - `npm run logs` - ログを表示
  - `npm run deploy` - 本番デプロイ

---

## 今後の拡張予定（Future Releases）

### v1.1.0 (計画中)

- [ ] 添付ファイルタイプのフィルタ機能拡張（PDF + 画像など）
- [ ] バッチ処理の自動スケジューリング
- [ ] Google Cloud Scheduler との連携
- [ ] 処理ログの永続化・ダウンロード機能

### v1.2.0 (計画中)

- [ ] 高度な検索フィルタ（AND/OR 複合条件）
- [ ] ファイル名テンプレート機能
- [ ] メールアーカイブ機能
- [ ] Slack / Teams 連携通知

### v2.0.0 (計画中)

- [ ] Web UI（Google Apps Script Webapp）の実装
- [ ] Google Cloud Functions との連携
- [ ] BigQuery へのログ出力

---

## セマンティックバージョニング

このプロジェクトは[セマンティックバージョニング](https://semver.org/ja/)に従います。

バージョンフォーマット：`MAJOR.MINOR.PATCH`

- **MAJOR** - 互換性がない変更
- **MINOR** - 機能追加（後方互換性あり）
- **PATCH** - バグ修正

---

## コントリビューション

このプロジェクトへの改善提案やバグ報告は、GitHub Issues にお願いします。

---
