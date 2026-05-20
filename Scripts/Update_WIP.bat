@echo off
chcp 65001 > nul
echo ============================================================
echo   Tubex - WIP Updater
echo   (Manual step - paste Aurangzeb's WIP message when prompted)
echo ============================================================
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
