@echo off
chcp 65001 > nul
echo ======================================================
echo  スーパー耐久 ラップビューア セットアップ
echo ======================================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup_lap_viewer.ps1"

echo.
pause
