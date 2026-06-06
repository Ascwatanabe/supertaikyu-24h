@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo スーパー耐久 24時間 タイミング取得を開始します...
echo 車両一覧: %~dp0data\index.html
echo ドライバー一覧: %~dp0data\drivers.html
echo ブラウザが自動で開きます。終了は Ctrl+C です。
echo.

python timing_fetcher.py --history --html --quiet --open-browser
if errorlevel 1 (
    echo.
    echo エラーが発生しました。Python と requests がインストールされているか確認してください。
    echo   pip install -r requirements.txt
    pause
)
