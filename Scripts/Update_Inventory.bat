@echo off
chcp 65001 > nul
echo ============================================================
echo   Tubex - Inventory Updater
echo ============================================================
echo.
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
python update_inventory.py
echo.
echo Press any key to close...
pause >nul
