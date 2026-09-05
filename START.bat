@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

title ГНС — запуск приложения
set "APP_URL=http://127.0.0.1:8765"
set "APP_PYTHON=%~dp0.venv\Scripts\python.exe"
set "APP_FAST_MODELS=%LOCALAPPDATA%\GNSLetterAutomation\models\tessdata_fast"
set "APP_BEST_MODELS=%LOCALAPPDATA%\GNSLetterAutomation\models\tessdata_best"
rem Test mode: all outgoing mail uses the approved Gmail recipient.
set "GNS_OUTLOOK_ALLOW_TEST_SEND=true"

rem Если приложение уже работает, просто открыть его и не запускать второй сервер.
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "try { $r = Invoke-WebRequest -UseBasicParsing -Uri '%APP_URL%/health' -TimeoutSec 2; if ($r.StatusCode -eq 200) { Start-Process '%APP_URL%'; exit 0 } } catch {}; exit 1" >nul 2>nul
if not errorlevel 1 exit /b 0

rem Обычный запуск ничего не устанавливает. Подготовка выполняется только
rem при первом запуске или если обязательные компоненты действительно пропали.
call :validate_runtime
if errorlevel 1 goto setup
goto launch

:setup
echo Первый запуск: подготавливаю приложение. Это выполняется один раз.
echo.

if not exist "%APP_PYTHON%" (
    where py >nul 2>nul
    if errorlevel 1 goto no_python
    py -3.12 --version >nul 2>nul
    if errorlevel 1 goto no_python
    py -3.12 -m venv .venv
    if errorlevel 1 goto setup_error
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\setup.ps1"
if errorlevel 1 goto setup_error
call :validate_runtime
if errorlevel 1 goto setup_error

:launch
echo Запускаю приложение: %APP_URL%
echo Окно браузера откроется автоматически.
echo Чтобы остановить приложение, закройте это окно.
echo.

"%APP_PYTHON%" -m gns_app.launcher
set "APP_EXIT=%ERRORLEVEL%"

rem Код 2 означает, что приложение уже было запущено; launcher уже открыл браузер.
if "%APP_EXIT%"=="0" exit /b 0
if "%APP_EXIT%"=="2" exit /b 0

echo.
echo Не удалось запустить приложение. Код ошибки: %APP_EXIT%
pause
exit /b %APP_EXIT%

:no_python
echo.
echo Нужен Python 3.12 (64-bit). Установите его и снова откройте START.bat.
echo https://www.python.org/downloads/release/python-3120/
pause
exit /b 1

:setup_error
echo.
echo Не удалось подготовить приложение. Проверьте интернет и повторите запуск.
pause
exit /b 1

:validate_runtime
if not exist "%APP_PYTHON%" exit /b 1
"%APP_PYTHON%" -c "import struct, sys; raise SystemExit(0 if sys.version_info[:2] == (3, 12) and struct.calcsize('P') * 8 == 64 else 1)" >nul 2>nul
if errorlevel 1 exit /b 1
"%APP_PYTHON%" -c "import gns_app, tesserocr" >nul 2>nul
if errorlevel 1 exit /b 1
for %%F in (
    "%APP_FAST_MODELS%\rus.traineddata"
    "%APP_FAST_MODELS%\kir.traineddata"
    "%APP_FAST_MODELS%\osd.traineddata"
    "%APP_BEST_MODELS%\rus.traineddata"
    "%APP_BEST_MODELS%\kir.traineddata"
    "%APP_BEST_MODELS%\osd.traineddata"
) do (
    if not exist "%%~F" exit /b 1
    if %%~zF LEQ 0 exit /b 1
)
exit /b 0
