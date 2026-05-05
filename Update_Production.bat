@echo off
chcp 65001 > nul
echo ============================================================
echo   Alpha Containers - Production Log Updater
echo ============================================================
echo.
echo   Installing libraries (first time only)...
pip install pandas openpyxl >nul 2>nul
echo   Ready.
echo.
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
python update_production.py
echo.
echo Press any key to close...
pause >nul
