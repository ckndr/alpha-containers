@echo off
:: Removes the Windows Task Scheduler job created by Install_Hourly_Push.bat.
if not defined _KEEP_OPEN (
    set _KEEP_OPEN=1
    cmd /k "%~f0"
    exit
)
setlocal

set "TASK_NAME=Tubex Hourly Push"

echo.
echo  ===========================================================
echo   Tubex -- Remove Hourly GitHub Push
echo  ===========================================================
echo.

schtasks /Delete /F /TN "%TASK_NAME%"
if errorlevel 1 (
    echo  ERROR: Could not remove the scheduled task. It may not exist.
) else (
    echo  Hourly push task removed.
)

echo.
echo  ===========================================================
echo.
pause
endlocal
