@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PATCH=%~dp0"
set "TARGET=%~dp0"
if not exist "%TARGET%GNS-Portable.exe" set "TARGET=%~dp0..\"
for %%I in ("%TARGET%") do set "TARGET=%%~fI"

if not exist "%TARGET%\GNS-Portable.exe" goto wrong_place
if not exist "%PATCH%files\GNS-Portable.exe" goto broken_patch
if not exist "%PATCH%files\GNS-Worker.exe" goto broken_patch
if not exist "%PATCH%files\_internal\gns_app" goto broken_patch

:wait_until_closed
tasklist /FI "IMAGENAME eq GNS-Portable.exe" /NH | find /I "GNS-Portable.exe" >nul
if not errorlevel 1 goto app_running
tasklist /FI "IMAGENAME eq GNS-Worker.exe" /NH | find /I "GNS-Worker.exe" >nul
if not errorlevel 1 goto app_running
goto install

:app_running
echo Close GNS-Portable, then press any key.
pause >nul
goto wait_until_closed

:install
set "BACKUP=%TARGET%\.update-backup"
if not exist "%BACKUP%" mkdir "%BACKUP%" >nul 2>&1
if not exist "%BACKUP%\_internal" mkdir "%BACKUP%\_internal" >nul 2>&1
if not exist "%BACKUP%\runtime" mkdir "%BACKUP%\runtime" >nul 2>&1

copy /Y "%TARGET%\GNS-Portable.exe" "%BACKUP%\GNS-Portable.exe" >nul || goto backup_failed
copy /Y "%TARGET%\GNS-Worker.exe" "%BACKUP%\GNS-Worker.exe" >nul || goto backup_failed
xcopy "%TARGET%\_internal\gns_app" "%BACKUP%\_internal\gns_app\" /E /I /Y /Q >nul || goto backup_failed
if exist "%TARGET%\runtime\gns.sqlite3" copy /Y "%TARGET%\runtime\gns.sqlite3" "%BACKUP%\runtime\gns.sqlite3.pre-update" >nul || goto backup_failed
if exist "%TARGET%\runtime\gns.sqlite3-wal" copy /Y "%TARGET%\runtime\gns.sqlite3-wal" "%BACKUP%\runtime\gns.sqlite3-wal.pre-update" >nul || goto backup_failed
if exist "%TARGET%\runtime\gns.sqlite3-shm" copy /Y "%TARGET%\runtime\gns.sqlite3-shm" "%BACKUP%\runtime\gns.sqlite3-shm.pre-update" >nul || goto backup_failed

copy /Y "%PATCH%files\GNS-Portable.exe" "%TARGET%\GNS-Portable.exe" >nul || goto rollback
copy /Y "%PATCH%files\GNS-Worker.exe" "%TARGET%\GNS-Worker.exe" >nul || goto rollback
xcopy "%PATCH%files\_internal\gns_app" "%TARGET%\_internal\gns_app\" /E /I /Y /Q >nul || goto rollback
if exist "%PATCH%files\START.bat" copy /Y "%PATCH%files\START.bat" "%TARGET%\START.bat" >nul
if exist "%PATCH%files\README.txt" copy /Y "%PATCH%files\README.txt" "%TARGET%\README.txt" >nul

echo Update installed. History and work files were preserved.
if /I "%GNS_UPDATE_NO_START%"=="1" exit /b 0
start "" "%TARGET%\GNS-Portable.exe"
exit /b 0

:rollback
echo Update failed. Restoring the previous application files...
copy /Y "%BACKUP%\GNS-Portable.exe" "%TARGET%\GNS-Portable.exe" >nul
copy /Y "%BACKUP%\GNS-Worker.exe" "%TARGET%\GNS-Worker.exe" >nul
xcopy "%BACKUP%\_internal\gns_app" "%TARGET%\_internal\gns_app\" /E /I /Y /Q >nul
echo Previous version restored. History was not changed.
pause
exit /b 5

:wrong_place
echo Put the GNS-Update folder inside the existing GNS-Portable folder.
pause
exit /b 3

:broken_patch
echo Update files are incomplete. Extract the ZIP again.
pause
exit /b 4

:backup_failed
echo Cannot create a backup. The update was not installed.
pause
exit /b 6
