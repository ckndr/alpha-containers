@echo off
chcp 65001 > nul
echo ============================================================
echo   Alpha Containers - Update Production from WhatsApp Images
echo ============================================================
echo.
echo   Installing new Google GenAI library...
pip uninstall -y google-generativeai
pip install pandas openpyxl pillow google-genai
echo.
echo   Ready.
echo.
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
python update_from_images.py
echo.
echo Press any key to close...
pause >nul