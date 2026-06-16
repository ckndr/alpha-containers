@echo off
if not defined _KEEP_OPEN (
    set _KEEP_OPEN=1
    cmd /k "%~f0"
    exit
)
cd /d "%~dp0"
echo.
python daily.py %*
