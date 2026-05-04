@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo  ╔═══════════════════════════════════════════════════════╗
echo  ║   Alpha Containers — Daily Update + Deploy           ║
echo  ╚═══════════════════════════════════════════════════════╝
echo.

:: ── STEP 1: Run all existing data scripts ───────────────────
echo [1/4] Updating production data...
python update_production.py
if errorlevel 1 ( echo  ✗ update_production.py failed & goto :error )

echo [2/4] Updating inventory...
python update_inventory.py
if errorlevel 1 ( echo  ✗ update_inventory.py failed & goto :error )

echo [3/4] Updating HTML dashboard...
python update_html.py
if errorlevel 1 ( echo  ✗ update_html.py failed & goto :error )

:: ── STEP 2: Push to GitHub ───────────────────────────────────
echo.
echo [4/4] Pushing to GitHub Pages...

:: Check git is available
where git >nul 2>&1
if errorlevel 1 (
    echo  ✗ Git not found. See SETUP_GUIDE.md Step 3.
    goto :error
)

:: Check we are inside a git repo
git rev-parse --git-dir >nul 2>&1
if errorlevel 1 (
    echo  ✗ This folder is not a git repo yet. See SETUP_GUIDE.md Step 4.
    goto :error
)

:: Stage only the HTML file (never commit the Excel)
git add AlphaContainers_App.html

:: Only commit if there are actual changes
git diff --cached --quiet
if errorlevel 1 (
    for /f "tokens=1-3 delims=/ " %%a in ("%date%") do set TODAY=%%a-%%b-%%c
    git commit -m "Daily update !TODAY!"
    git push origin main
    if errorlevel 1 ( echo  ✗ git push failed. Check internet/credentials. & goto :error )
    echo  ✓ Deployed to GitHub Pages
) else (
    echo  ✓ No changes since last deploy — skipping push
)

echo.
echo  ════════════════════════════════════════════════════════
echo  ✓ All done. Boss's dashboard is live.
echo  ════════════════════════════════════════════════════════
echo.
timeout /t 4 >nul
goto :end

:error
echo.
echo  ════════════════════════════════════════════════════════
echo  ✗ Something failed. Check the error message above.
echo  ════════════════════════════════════════════════════════
pause

:end
endlocal
