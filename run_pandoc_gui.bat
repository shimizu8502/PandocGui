@echo off
:: ===============================================
:: Pandoc GUI 起動スクリプト
:: ===============================================

:: コードページをUTF-8に設定（文字化け対策）
chcp 65001 >nul 2>&1

echo.
echo ====================================
echo  Pandoc GUI - ファイル変換ツール
echo ====================================
echo.

:: 現在のディレクトリをスクリプトのディレクトリに変更
cd /d "%~dp0"

:: Pythonがインストールされているかチェック
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [エラー] Pythonがインストールされていないか、パスが通っていません。
    echo.
    echo 解決方法:
    echo 1. Python 3.6以降をインストールしてください
    echo 2. インストール時に「Add Python to PATH」にチェックを入れてください
    echo 3. https://www.python.org/downloads/ からダウンロードできます
    echo.
    pause
    exit /b 1
)

:: pandoc_gui.pyファイルが存在するかチェック
if not exist "pandoc_gui.py" (
    echo [エラー] pandoc_gui.py ファイルが見つかりません。
    echo このバッチファイルと同じフォルダに pandoc_gui.py を配置してください。
    echo.
    pause
    exit /b 1
)

echo Pandoc GUI を起動しています...
echo.

:: Pythonスクリプトを実行
python pandoc_gui.py

:: エラーが発生した場合の処理
if %errorlevel% neq 0 (
    echo.
    echo [エラー] アプリケーションの実行中にエラーが発生しました。
    echo エラーコード: %errorlevel%
    echo.
    echo 考えられる原因:
    echo - 必要なPythonライブラリが不足している
    echo - ファイルのアクセス権限の問題
    echo - Pythonのバージョンが古い
    echo.
    pause
)

exit /b %errorlevel% 