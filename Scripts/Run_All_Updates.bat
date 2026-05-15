@echo off
:: If not already inside a persistent window, relaunch inside cmd /k so the
:: window stays open after the script finishes. Plain pause is unreliable when
:: double-clicking from Explorer in some Windows configurations.
if not defined _KEEP_OPEN (
    set _KEEP_OPEN=1
    cmd /k "%~f0"
    exit
)
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo  ===========================================================
echo   Alpha Containers -- Daily Update + Deploy
echo  ===========================================================
echo.
echo  Sequence: Production -^> Inventory -^> Dispatch -^> HTML/GitHub
echo  Skipped (manual): MRP, WIP, update_from_images
echo.

set FAIL_COUNT=0
set FAIL_LIST=

:: ── STEP 1: Production Log + FG Stock ────────────────────────────────────────
:: Must run first — populates Production_Log which HTML and Dashboard read.
echo [1/4] Updating Production Log + FG Stock...
echo       (reads Production.xlsx ^-^> writes Production_Log + FG Stock)
python update_production.py
if errorlevel 1 (
    set /a FAIL_COUNT+=1
    set "FAIL_LIST=!FAIL_LIST! | [1] update_production"
    echo.
    echo  !! STEP 1 FAILED — Inventory and Dispatch will still run.
    echo  !! HTML will be generated from whatever is currently in Excel.
) else (
    echo  ^^ Step 1 OK.
)
echo.

:: ── STEP 2: Inventory ────────────────────────────────────────────────────────
:: Independent of Production — reads inventory.xls, writes Inventory sheet.
echo [2/4] Updating Inventory...
echo       (reads inventory.xls ^-^> writes Inventory sheet)
python update_inventory.py
if errorlevel 1 (
    set /a FAIL_COUNT+=1
    set "FAIL_LIST=!FAIL_LIST! | [2] update_inventory"
    echo.
    echo  !! STEP 2 FAILED — Inventory sheet not updated.
) else (
    echo  ^^ Step 2 OK.
)
echo.

:: ── STEP 3: Dispatch ─────────────────────────────────────────────────────────
:: Reads dispatch.xls + dispatch_pet.xls, writes Dashboard col K.
:: Must run before HTML (HTML reads col K for dispatch totals).
echo [3/4] Updating Dispatch...
echo       (reads dispatch.xls + dispatch_pet.xls ^-^> writes Dashboard col K)
python update_dispatch.py
if errorlevel 1 (
    set /a FAIL_COUNT+=1
    set "FAIL_LIST=!FAIL_LIST! | [3] update_dispatch"
    echo.
    echo  !! STEP 3 FAILED — Dashboard dispatch column not updated.
) else (
    echo  ^^ Step 3 OK.
)
echo.

:: ── STEP 4: HTML Dashboard ───────────────────────────────────────────────────
:: Must run last — reads Production_Log, Dashboard (col K), Catalog, BOM.
echo [4/4] Generating HTML Dashboard...
echo       (reads Excel ^-^> writes AlphaContainers_App.html)
python update_html.py
if errorlevel 1 (
    set /a FAIL_COUNT+=1
    set "FAIL_LIST=!FAIL_LIST! | [4] update_html"
    echo.
    echo  !! STEP 4 FAILED — HTML not updated. Check error above.
) else (
    echo  ^^ Step 4 OK.
)
echo.

:: ── GIT PUSH ─────────────────────────────────────────────────────────────────
:: Back up the entire repository. Always attempt even if steps above partially failed,
:: so all generated files and project changes are saved to GitHub.
echo [Git] Backing up full repository to GitHub...

where git >nul 2>&1
if errorlevel 1 (
    set /a FAIL_COUNT+=1
    set "FAIL_LIST=!FAIL_LIST! | [Git] git.exe not found"
    echo  !! Git not installed or not in PATH. See SETUP_GUIDE for Step 3.
    goto :summary
)

git -C "%~dp0.." rev-parse --git-dir >nul 2>&1
if errorlevel 1 (
    set /a FAIL_COUNT+=1
    set "FAIL_LIST=!FAIL_LIST! | [Git] Not a repo yet"
    echo  !! AlphaContainers folder is not a git repo. See SETUP_GUIDE Step 4.
    goto :summary
)

:: Stage every change in the root repository, including new, modified, and deleted files.
git -C "%~dp0.." add -A

:: Only commit if there are staged changes
git -C "%~dp0.." diff --cached --quiet
if errorlevel 1 (
    git -C "%~dp0.." commit -m "Full backup %date% %time:~0,8%"
    git -C "%~dp0.." push origin main
    if errorlevel 1 (
        set /a FAIL_COUNT+=1
        set "FAIL_LIST=!FAIL_LIST! | [Git] push failed (check internet/credentials)"
        echo  !! Git push failed. Check internet connection and GitHub credentials.
    ) else (
        echo  ^^ Full repository backup pushed to main.
    )
) else (
    echo  ^^ No repository changes since last backup - skipping push.
)

:summary
echo.
echo  ===========================================================
if !FAIL_COUNT! == 0 (
    echo   ALL STEPS COMPLETED SUCCESSFULLY.
    echo.
    echo   Next: Open Excel and press Ctrl+Shift+F9 to recalculate.
    echo   Then: Check dashboard is live on GitHub Pages.
) else (
    echo   COMPLETED WITH !FAIL_COUNT! FAILURE(S):
    echo   !FAIL_LIST!
    echo.
    echo   Scroll up to see the error output for each failed step.
    echo   Steps that succeeded still saved their changes to Excel.
)
echo  ===========================================================
echo.
pause
