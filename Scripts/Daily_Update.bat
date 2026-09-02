@echo off
if not defined _KEEP_OPEN (
    set _KEEP_OPEN=1
    cmd /k "%~f0"
    exit
)
cd /d "%~dp0"
:: If running from OneDrive mirror, redirect to master repository on D:\Alpha
if /i "%CD%"=="C:\Users\HP\OneDrive\Alpha\Scripts" (
    echo [NOTICE] Running from OneDrive mirror. Switching to D:\Alpha master repository...
    cd /d "D:\Alpha\Scripts"
)
echo.
python daily.py %*
