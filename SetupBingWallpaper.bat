@echo off
cls

echo Bing Daily Wallpaper Setup Tool
echo Please ensure BingDailyWallpaper.ps1 and SilentRun.vbs are in the current directory
echo.

set ScriptPath=%~dp0BingDailyWallpaper.ps1
set VBSPath=%~dp0SilentRun.vbs

:start
echo Please select run mode:
echo 1. Run at startup
echo 2. Run daily at 8:10 AM
set /p choice=Enter 1 or 2: 

if "%choice%"=="1" (
    schtasks /create /tn "BingDailyWallpaper_Startup" /tr "wscript.exe \"%VBSPath%\" \"%ScriptPath%\"" /sc onlogon /rl highest /f /delay 0000:30
    echo Startup task has been set
) else if "%choice%"=="2" (
    schtasks /create /tn "BingDailyWallpaper_0810" /tr "wscript.exe \"%VBSPath%\" \"%ScriptPath%\"" /sc daily /st 08:10 /rl highest /f
    echo Daily 8:10 AM task has been set
) else (
    echo Invalid selection, please try again
    goto start
)

echo Setup completed. Press any key to exit
pause