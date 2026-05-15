@echo off
:: Creates/updates a Windows Task Scheduler job that runs Push.bat every hour.
if not defined _KEEP_OPEN (
    set _KEEP_OPEN=1
    cmd /k "%~f0"
    exit
)
setlocal

set "TASK_NAME=Alpha Containers Hourly Push"
set "PUSH_BAT=%~dp0Push.bat"

echo.
echo  ===========================================================
echo   Alpha Containers -- Install Hourly GitHub Push
echo  ===========================================================
echo.
echo  This will create a Windows scheduled task:
echo    %TASK_NAME%
echo.
echo  It will run every hour while this PC is on and logged in.
echo  Logs will be written to:
echo    %~dp0..\Logs\hourly_push.log
echo.

where schtasks >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Windows Task Scheduler command not found.
    goto :done
)

schtasks /Create /F /TN "%TASK_NAME%" /SC HOURLY /MO 1 /TR "\"%PUSH_BAT%\" /auto"
if errorlevel 1 (
    echo.
    echo  ERROR: Could not create the scheduled task.
    echo  Try right-clicking this file and choosing Run as administrator.
    goto :done
)

echo.
echo  Hourly push task installed.
echo.
echo  Tip: run Scripts\Push.bat manually once first, so GitHub credentials
echo  are already saved before the scheduled task runs unattended.

:done
echo.
echo  ===========================================================
echo.
pause
endlocal
