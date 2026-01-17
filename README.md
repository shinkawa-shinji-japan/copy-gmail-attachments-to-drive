# Gmail添付ファイルをGoogleドライブにコピーするツール

## 概要

Gmailの指定期間内にある**PDF添付ファイル**を、Google Drive の指定フォルダに一括コピーするGoogle Apps Scriptツールです。

### 主な機能

✅ **期間指定検索** - 指定日付範囲のメールを検索  
✅ **キーワードフィルタ** - メールタイトルで検索  
✅ **PDF抽出** - PDF添付ファイルのみを自動抽出  
✅ **ワンクリック保存** - チェックボックスで保存対象を選択して実行  
✅ **フォルダ自動作成** - 保存先フォルダが無ければ自動作成  
✅ **重複回避** - 同名ファイルがあればスキップ  
✅ **直リンク生成** - コピー後、Google Drive ファイルへのリンク自動生成

---

## クイックスタート

### 1. セットアップ

```bash
# リポジトリをクローン
git clone https://github.com/shinkawa-shinji-japan/copy-gmail-attachments-to-drive.git
cd copy-gmail-attachments-to-drive

# Google Apps Script プロジェクトを作成（詳細は SETUP_GUIDE.md を参照）
clasp login
clasp create --type sheets --title "Gmail Attachments to Drive"
```

### 2. 初期化

1. Google Sheets を開く
2. メニュー「Gmail添付ファイル」→「初期化（シート作成）」を実行

### 3. 検索条件を設定

「検索条件」シートで以下を入力：

- 期間（開始日、終了日）
- フィルタ文字列（オプション）
- 保存先フォルダパス

### 4. メールを検索

メニュー「Gmail添付ファイル」→「検索して一覧表示」を実行

### 5. ファイルをコピー

チェックボックスで対象を選択し、「選択対象をGoogleドライブにコピー」を実行

詳細は [SETUP_GUIDE.md](SETUP_GUIDE.md) を参照してください。

---

## ドキュメント

| ファイル                                         | 内容                         |
| ------------------------------------------------ | ---------------------------- |
| [REQUIREMENTS.md](REQUIREMENTS.md)               | 要件定義・機能仕様           |
| [IMPLEMENTATION_PLAN.md](IMPLEMENTATION_PLAN.md) | 実装方針・技術仕様           |
| [SETUP_GUIDE.md](SETUP_GUIDE.md)                 | セットアップ・デプロイガイド |

---

## ファイル構成

```
copy-gmail-attachments-to-drive/
├── src/
│   ├── main.gs          # メイン処理・UI制御
│   ├── gmail.gs         # Gmail API ラッパー関数
│   ├── drive.gs         # Google Drive API ラッパー関数
│   ├── sheet.gs         # Google Sheets ラッパー関数
│   └── utils.gs         # ユーティリティ関数
├── appsscript.json      # Google Apps Script マニフェスト
├── package.json         # Node.js / clasp 設定
├── README.md            # このファイル
├── REQUIREMENTS.md      # 要件定義
├── IMPLEMENTATION_PLAN.md  # 実装方針
└── SETUP_GUIDE.md       # セットアップガイド
```

---

## 使用技術

- **Google Apps Script** - メイン開発言語
- **Gmail API** - メール検索・添付ファイル取得
- **Google Drive API** - フォルダ・ファイル操作
- **Google Sheets API** - スプレッドシート操作
- **clasp** - ローカル開発・デプロイ

---

## 前提条件

- Google アカウント（Gmail、Google Drive アクセス可能）
- Node.js 18以上
- clasp がインストール済み

```bash
npm install -g @google/clasp
```

---

## コマンドリファレンス

```bash
# 認証
clasp login

# コードを Google Apps Script にプッシュ
npm run push

# Google Apps Script からコードをプル
npm run pull

# Google Apps Script エディタを開く
npm run open

# ログを表示
npm run logs

# 本番デプロイ
npm run deploy
```

---

## 使用例

### 例1：2026年1月の「請求書」PDF を `/請求書/2026年1月` に保存

**検索条件シートの入力：**

```
期間（開始日）: 2026-01-01
期間（終了日）: 2026-01-31
フィルタ文字列: 請求書
保存先フォルダパス: /請求書/2026年1月
```

**実行：**

1. 「検索して一覧表示」を実行
2. 検索結果を確認
3. 「選択対象をGoogleドライブにコピー」を実行

### 例2：複数キーワードで検索

**検索条件シートの入力：**

```
フィルタ文字列: 請求書, 領収書, 見積書
```

複数キーワードはカンマで区切ります（OR条件）

---

## トラブルシューティング

### メニューが表示されない

1. Google Apps Script エディタで「初期化（シート作成）」を実行
2. 権限ダイアログで「許可」をクリック
3. Google Sheets をリロード

### 検索結果が0件

- 日付範囲を確認
- フィルタ文字列を削除して再試行

### ファイルコピーが失敗

1. フォルダパスを確認（フォルダ名のスペースを含む正確な名称）
2. Google Drive の残容量を確認
3. 結果シートのエラーメッセージを確認

詳細は [SETUP_GUIDE.md](SETUP_GUIDE.md) のトラブルシューティングを参照してください。

---

## セキュリティ

このスクリプトは以下の権限を要求します：

- **Gmail（読み取り専用）** - メール検索と添付ファイル取得
- **Google Drive** - フォルダ・ファイル操作
- **Google Sheets** - スプレッドシート読み書き

すべてのアクセスはユーザーの明示的な認可に基づいています。

---

## ライセンス

MIT License

---

## 作成者

[GitHub: shinkawa-shinji-japan](https://github.com/shinkawa-shinji-japan)

---

## 更新履歴

### v1.0.0 (2026-01-17)

- 初版リリース
- 基本機能実装完了
  - メール検索・抽出
  - PDF フォルダ保存
  - ファイル名カスタマイズ
  - 重複ファイルのスキップ

---

## サポート

問題や機能リクエストは GitHub Issues で報告してください。
