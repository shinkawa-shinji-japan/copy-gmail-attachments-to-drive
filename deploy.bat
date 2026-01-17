@echo off
REM デプロイ自動化スクリプト（Windows用）

echo ================================
echo Gmail Attachments to Drive
echo デプロイスクリプト
echo ================================
echo.

REM Node.js と npm のバージョン確認
echo ^[OK^] Node.js と npm のバージョン確認
node --version
npm --version
echo.

REM clasp インストール確認
where clasp >nul 2>nul
if errorlevel 1 (
    echo [WARN] clasp がインストールされていません。インストール中...
    call npm install -g @google/clasp
)

echo ^[OK^] clasp バージョン確認
call clasp --version
echo.

REM プロジェクトディレクトリに移動
cd /d "%~dp0"

REM package.json の依存パッケージをインストール
echo [INFO] npm 依存パッケージをインストール中...
call npm install
echo.

REM ユーザーに選択肢を提示
echo 次の操作を選択してください:
echo 1) clasp login（Google アカウントで認証）
echo 2) npm run push（コードをプッシュ）
echo 3) npm run open（Google Apps Script エディタを開く）
echo 4) npm run logs（ログを表示）
echo 5) 終了
echo.

set /p choice="選択 (1-5): "

if "%choice%"=="1" (
    echo [INFO] clasp ログイン実行中...
    call clasp login
) else if "%choice%"=="2" (
    echo [INFO] コードをプッシュ中...
    call npm run push
) else if "%choice%"=="3" (
    echo [INFO] Google Apps Script エディタを開く中...
    call npm run open
) else if "%choice%"=="4" (
    echo [INFO] ログを表示中...
    call npm run logs
) else if "%choice%"=="5" (
    echo 終了します
    exit /b 0
) else (
    echo 無効な選択です
    exit /b 1
)

echo.
echo ^[OK^] 完了！
pause
