@echo off
if /i "%~1"=="/auto" set "_AUTO_PUSH=1"
:: If not already inside a persistent window, relaunch inside cmd /k so the
:: window stays open after the script finishes.
if not defined _KEEP_OPEN if not defined _AUTO_PUSH (
    set _KEEP_OPEN=1
    cmd /k "%~f0"
    exit
)
setlocal enabledelayedexpansion
cd /d "%~dp0.."

if /i "%~1"=="/auto" (
    if not exist "Logs" mkdir "Logs"
    >>"Logs\hourly_push.log" echo.
    >>"Logs\hourly_push.log" echo ===========================================================
    >>"Logs\hourly_push.log" echo Hourly push started: %date% %time%
    set "_AUTO_PUSH="
    call "%~f0" /run >>"Logs\hourly_push.log" 2>&1
    exit /b %errorlevel%
)

if /i "%~1"=="/run" goto :run_push

:run_push
echo.
echo  ===========================================================
echo   Alpha Containers -- Push Full Backup to GitHub
echo  ===========================================================
echo.

where git >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Git is not installed or not in PATH.
    goto :done
)

git rev-parse --git-dir >nul 2>&1
if errorlevel 1 (
    echo  ERROR: This folder is not a git repository.
    goto :done
)

echo [1/3] Staging every repository change...
git add -A
if errorlevel 1 (
    echo  ERROR: Git add failed.
    goto :done
)

echo [2/3] Checking for changes...
git diff --cached --quiet
if not errorlevel 1 (
    echo  No local changes to push.
    goto :done
)

echo [3/3] Committing and pushing to main...
git commit -m "Full backup %date% %time:~0,8%"
if errorlevel 1 (
    echo  ERROR: Commit failed.
    goto :done
)

git push origin main
if errorlevel 1 (
    echo  ERROR: Push failed. Check internet connection and GitHub credentials.
    goto :done
)

echo.
echo  Full backup pushed to GitHub main.

:done
echo.
echo  ===========================================================
echo.
if /i "%~1"=="/run" exit /b 0
pause
endlocal
