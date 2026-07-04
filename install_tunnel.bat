@echo off

chcp 65001 >nul

echo ========================================
echo  Cloudflare Tunnel インストーラー
echo ========================================
echo.

where cloudflared >nul 2>&1

if %errorlevel% == 0 (

  echo [OK] cloudflared はすでにインストールされています。

  cloudflared --version

  echo.

  echo start.bat を実行するとトンネルが自動で起動します。

  pause

  exit /b 0

)

echo cloudflared をインストールします (winget)...

echo.

winget install Cloudflare.cloudflared

if errorlevel 1 (

  echo.

  echo [エラー] winget でのインストールに失敗しました。

  echo 手動インストール: https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/downloads/

  pause

  exit /b 1

)

echo.

echo ========================================

echo  インストール完了！

echo ========================================

echo.

echo 次回 start.bat を実行すると自動でトンネルが起動し、

echo 外部公開URL（https://xxxx.trycloudflare.com）が表示されます。

echo そのURLをチームの方々に共有してください。

echo.

pause
