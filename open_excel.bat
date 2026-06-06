@echo off

chcp 65001 >nul

cd /d "%~dp0"

set "EXCEL_FILE=%~dp0data\timing_live.xlsx"

if not exist "%EXCEL_FILE%" (
    echo Excelファイルがまだありません。先に start.bat を実行してください。
    pause
    exit /b 1
)

start "" "%EXCEL_FILE%"
