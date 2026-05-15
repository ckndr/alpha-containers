@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo  ================================================
echo   Alpha Containers - Update HTML + Publish
echo   (Manual: runs HTML only, skips Production/Inventory/Dispatch)
echo  ================================================
echo.

echo [1/2] Generating HTML from Excel...
set PYTHONIOENCODING=utf-8
python update_html.py
if errorlevel 1 (
    echo.
    echo  ERROR: update_html.py failed. Check output above.
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

git -C "%~dp0.." rev-parse --git-dir >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Not a git repo. Follow SETUP_GUIDE Step 4.
    pause
    exit /b 1
)

git -C "%~dp0.." add AlphaContainers_App.html
if exist "%~dp0..\sw.js"         git -C "%~dp0.." add sw.js
if exist "%~dp0..\manifest.json" git -C "%~dp0.." add manifest.json
if exist "%~dp0..\icon-192.png"  git -C "%~dp0.." add icon-192.png
if exist "%~dp0..\icon-512.png"  git -C "%~dp0.." add icon-512.png

git -C "%~dp0.." diff --cached --quiet
if errorlevel 1 (
    git -C "%~dp0.." commit -m "Dashboard update %date% %time:~0,5%"
    git -C "%~dp0.." push origin main
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
echo   Done. Dashboard is updated and live.
echo  ================================================
echo.
pause
endlocal
