@echo off
:: If not already inside a persistent window, relaunch inside cmd /k so the
:: window stays open after the script finishes.
if not defined _KEEP_OPEN (
    set _KEEP_OPEN=1
    cmd /k "%~f0"
    exit
)
setlocal
cd /d "%~dp0.."

echo.
echo  ===========================================================
echo   Tubex -- Pull Latest Files from GitHub
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

echo [1/3] Checking for local uncommitted changes...
git diff --quiet
if errorlevel 1 (
    echo.
    echo  ERROR: You have local file changes that are not committed yet.
    echo  Run Scripts\Push.bat first, or commit/stash those changes before pulling.
    goto :done
)

git diff --cached --quiet
if errorlevel 1 (
    echo.
    echo  ERROR: You have staged changes that are not committed yet.
    echo  Run Scripts\Push.bat first, or commit/stash those changes before pulling.
    goto :done
)

echo [2/3] Downloading latest GitHub history...
git fetch origin main
if errorlevel 1 (
    echo  ERROR: Fetch failed. Check internet connection and GitHub credentials.
    goto :done
)

echo [3/3] Updating this folder from origin/main...
git pull --ff-only origin main
if errorlevel 1 (
    echo.
    echo  ERROR: Pull failed. Your local and GitHub histories may have diverged.
    echo  Ask Codex to help resolve the Git sync issue.
    goto :done
)

echo.
echo  This folder is now updated from GitHub main.

:done
echo.
echo  ===========================================================
echo.
pause
endlocal
