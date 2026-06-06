@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ローカルサーバーで車両一覧を開きます...
echo 終了は Ctrl+C です。
echo.

python timing_fetcher.py --serve-browser --serve-page index
