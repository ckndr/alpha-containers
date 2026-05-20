@echo off
chcp 65001 > nul
echo ============================================================
echo   Tubex - Production Log Updater
echo ============================================================
echo.
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
python update_production.py
echo.
echo Press any key to close...
pause >nul
