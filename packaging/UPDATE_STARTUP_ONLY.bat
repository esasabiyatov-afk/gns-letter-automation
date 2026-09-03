@echo off
setlocal
cd /d "%~dp0"

set "TARGET=%~dp0.."
if not exist "%TARGET%\GNS-Portable.exe" goto wrong_place
if not exist "%~dp0GNS-Portable.exe" goto broken_patch

tasklist /FI "IMAGENAME eq GNS-Portable.exe" /NH | find /I "GNS-Portable.exe" >nul
if not errorlevel 1 goto app_running

copy /Y "%~dp0GNS-Portable.exe" "%TARGET%\GNS-Portable.exe" >nul || goto copy_failed
echo Update installed successfully.
start "" "%TARGET%\GNS-Portable.exe"
exit /b 0

:app_running
echo Close GNS-Portable and run UPDATE.bat again.
pause
exit /b 2

:wrong_place
echo Copy this hotfix folder inside the old GNS-Portable folder.
pause
exit /b 3

:broken_patch
echo Update file is missing. Extract the ZIP again.
pause
exit /b 4

:copy_failed
echo Update failed. Close the application and try again.
pause
exit /b 5
