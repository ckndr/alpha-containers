@echo off
chcp 65001 > nul
echo ============================================================
echo   Alpha Containers - Dispatch Updater
echo ============================================================
echo.
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
python update_dispatch.py
echo.
echo Press any key to close...
pause >nul
