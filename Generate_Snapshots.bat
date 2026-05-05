@echo off
chcp 65001 > nul
echo ============================================================
echo   Alpha Containers - Snapshot Generator
echo ============================================================
echo.
echo   Installing libraries (first time only)...
pip install openpyxl Pillow >nul 2>nul
echo   Ready.
echo.
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
python generate_snapshots.py
echo.
echo Press any key to close...
pause >nul
