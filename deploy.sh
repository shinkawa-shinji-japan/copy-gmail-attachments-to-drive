#!/bin/bash
# デプロイ自動化スクリプト

echo "================================"
echo "Gmail Attachments to Drive"
echo "デプロイスクリプト"
echo "================================"
echo ""

# Node.js と npm のバージョン確認
echo "✓ Node.js と npm のバージョン確認"
node --version
npm --version
echo ""

# clasp インストール確認
if ! command -v clasp &> /dev/null; then
    echo "⚠ clasp がインストールされていません。インストール中..."
    npm install -g @google/clasp
fi

echo "✓ clasp バージョン確認"
clasp --version
echo ""

# プロジェクトディレクトリに移動
cd "$(dirname "$0")" || exit

# package.json の依存パッケージをインストール
echo "📦 npm 依存パッケージをインストール中..."
npm install
echo ""

# ユーザーに選択肢を提示
echo "次の操作を選択してください："
echo "1) clasp login（Google アカウントで認証）"
echo "2) npm run push（コードをプッシュ）"
echo "3) npm run open（Google Apps Script エディタを開く）"
echo "4) npm run logs（ログを表示）"
echo "5) 終了"
echo ""

read -p "選択 (1-5): " choice

case $choice in
  1)
    echo "🔑 clasp ログイン実行中..."
    clasp login
    ;;
  2)
    echo "📤 コードをプッシュ中..."
    npm run push
    ;;
  3)
    echo "🌐 Google Apps Script エディタを開く中..."
    npm run open
    ;;
  4)
    echo "📋 ログを表示中..."
    npm run logs
    ;;
  5)
    echo "終了します"
    exit 0
    ;;
  *)
    echo "無効な選択です"
    exit 1
    ;;
esac

echo ""
echo "✓ 完了！"
