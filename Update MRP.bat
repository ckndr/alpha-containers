@echo off
echo ============================================================
echo   Alpha Containers - MRP Updater
echo ============================================================
echo.
echo   Installing libraries (first time only)...
pip install openpyxl >nul 2>&1
echo   Ready.
echo.
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
python update_mrp.py
echo.
echo Press any key to close...
pause >nul
