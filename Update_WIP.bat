@echo off
chcp 65001 > nul
echo ============================================================
echo   Alpha Containers - WIP Updater
echo ============================================================
echo.
echo   Installing libraries (first time only)...
pip install openpyxl >nul 2>nul
echo   Ready.
echo.
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8

if "%~1"=="" (
    python update_wip.py
) else (
    python update_wip.py %*
)

echo.
echo Press any key to close...
pause >nul
