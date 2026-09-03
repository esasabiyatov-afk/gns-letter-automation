@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

if not exist "GNS-Portable.exe" (
  echo Не найден GNS-Portable.exe. Распакуйте архив полностью.
  pause
  exit /b 1
)

rem Python и интернет для запуска не требуются. START закрывается сразу,
rem само приложение и его служебные процессы работают без окна терминала.
start "" "%~dp0GNS-Portable.exe"
exit /b 0
