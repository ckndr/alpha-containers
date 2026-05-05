@echo off
chcp 65001 > nul
echo ============================================================
echo   Alpha Containers - Inventory Updater
echo ============================================================
echo.
echo   Installing libraries (first time only)...
pip install pandas openpyxl xlrd >nul 2>nul
echo   Ready.
echo.
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
python update_inventory.py
echo.
echo Press any key to close...
pause >nul
