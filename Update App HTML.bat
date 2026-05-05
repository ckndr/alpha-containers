@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo  ================================================
echo   Alpha Containers - Update Dashboard + Publish
echo  ================================================
echo.

:: Check xlrd
python -c "import xlrd" >nul 2>&1
if errorlevel 1 (
    echo  Installing xlrd...
    pip install xlrd --quiet 2>nul || pip install xlrd --break-system-packages --quiet 2>nul
)

echo [1/2] Generating HTML from Excel...
python update_html.py
if errorlevel 1 (
    echo.
    echo  ERROR: update_html.py failed. Check that:
    echo    - Excel file is closed (not open in Excel)
    echo    - AlphaContainers_v*.xlsx exists in this folder
    echo    - Python is installed correctly
    echo.
    pause
    exit /b 1
)

echo.
echo [2/2] Publishing to GitHub...

where git >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Git not found. Install from https://git-scm.com
    pause
    exit /b 1
)

git rev-parse --git-dir >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Not a git repo. Follow SETUP_GUIDE.txt Step 4.
    pause
    exit /b 1
)

git add AlphaContainers_App.html
:: Also stage icon/manifest/sw files if they exist
if exist manifest.json git add manifest.json
if exist sw.js git add sw.js
if exist icon-192.png git add icon-192.png
if exist icon-512.png git add icon-512.png
git diff --cached --quiet
if errorlevel 1 (
    git commit -m "Dashboard update %date% %time:~0,5%"
    git push origin HEAD
    if errorlevel 1 (
        echo.
        echo  ERROR: Push failed. Check internet connection and GitHub credentials.
        pause
        exit /b 1
    )
    echo  Published! Live in ~60 seconds.
) else (
    echo  No changes since last publish - skipping push.
)

echo.
echo  ================================================
echo   Done. Boss dashboard is updated and live.
echo  ================================================
echo.
pause
endlocal
