@echo off

chcp 65001 >nul

cd /d "%~dp0"



call "%~dp0stop.bat"



echo.

echo AndLegal Racing タイミング取得を開始します...

echo 車両一覧: http://127.0.0.1:8765/index.html

echo ドライバー一覧: http://127.0.0.1:8765/drivers.html

echo Excel出力: data\timing_live.xlsx

echo ブラウザが自動で開きます。終了は Ctrl+C です。

echo.



where cloudflared >nul 2>&1 && (

  echo.

  echo [トンネル] cloudflared を起動しています...

  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0_tunnel_helper.ps1"

)



python timing_fetcher.py --history --excel --html --quiet --open-browser

if errorlevel 1 (

    echo.

    echo エラーが発生しました。Python と requests がインストールされているか確認してください。

    echo   pip install -r requirements.txt

    pause

)
