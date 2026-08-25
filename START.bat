@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ================================================================
echo   GNS Letter Automation - setup and start
echo ================================================================
echo.
echo Working folder: %cd%
echo.

REM ---------------------------------------------------------------
REM 1. Find Python 3.12 (required exactly - the tesserocr wheel is
REM    built for cp312; newer versions like 3.14 will NOT work)
REM ---------------------------------------------------------------
where py >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python launcher "py" was not found.
    echo Install Python 3.12 ^(64-bit^) from:
    echo   https://www.python.org/downloads/release/python-3120/
    echo During setup, make sure to check "Add python.exe to PATH".
    echo.
    pause
    exit /b 1
)

py -3.12 --version >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python 3.12 is not installed ^(other versions, e.g. 3.14, will NOT work^).
    echo Download the "Windows installer ^(64-bit^)" here:
    echo   https://www.python.org/downloads/release/python-3120/
    echo During setup, check "Add python.exe to PATH", then run this file again.
    echo.
    pause
    exit /b 1
)

echo [OK] Python 3.12 found.
echo.

REM ---------------------------------------------------------------
REM 2. Create the virtual environment if it does not exist yet
REM ---------------------------------------------------------------
set "PYEXE=.venv\Scripts\python.exe"

if not exist "%PYEXE%" (
    echo Creating virtual environment ^(.venv^)...
    py -3.12 -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Failed to create the virtual environment.
        pause
        exit /b 1
    )
    echo [OK] Virtual environment created.
) else (
    echo [OK] Virtual environment already exists.
)
echo.

REM ---------------------------------------------------------------
REM 3. Install project dependencies
REM ---------------------------------------------------------------
echo Installing project dependencies ^(this can take a few minutes^)...
"%PYEXE%" -m pip install --upgrade pip --quiet
if errorlevel 1 (
    echo [ERROR] Failed to upgrade pip. Check your internet connection.
    pause
    exit /b 1
)

"%PYEXE%" -m pip install -e ".[dev]" --quiet
if errorlevel 1 (
    echo [ERROR] Failed to install project dependencies.
    pause
    exit /b 1
)
echo [OK] Dependencies installed.
echo.

REM ---------------------------------------------------------------
REM 4. Check tesserocr, install the prebuilt Windows wheel if needed
REM ---------------------------------------------------------------
"%PYEXE%" -c "import tesserocr" >nul 2>nul
if errorlevel 1 (
    echo tesserocr not found, installing the prebuilt Windows wheel...
    "%PYEXE%" -m pip install "https://github.com/simonflueckiger/tesserocr-windows_build/releases/download/tesserocr-v2.10.0-tesseract-5.5.2/tesserocr-2.10.0-cp312-cp312-win_amd64.whl"
    if errorlevel 1 (
        echo [ERROR] Failed to install tesserocr.
        pause
        exit /b 1
    )
    echo [OK] tesserocr installed.
) else (
    echo [OK] tesserocr already installed.
)
echo.

REM ---------------------------------------------------------------
REM 5. Lay out the OCR language models (rus/kir/osd)
REM    First try the bundled models/ folder from this archive - it
REM    is faster and needs no internet. Anything still missing gets
REM    downloaded automatically.
REM ---------------------------------------------------------------
set "MODEL_ROOT=%LOCALAPPDATA%\GNSLetterAutomation\models"
set "FAST_DST=%MODEL_ROOT%\tessdata_fast"
set "BEST_DST=%MODEL_ROOT%\tessdata_best"

if not exist "%FAST_DST%" mkdir "%FAST_DST%" >nul 2>nul
if not exist "%BEST_DST%" mkdir "%BEST_DST%" >nul 2>nul

if not exist "%FAST_DST%\rus.traineddata" (
    echo Copying OCR language models...
    if exist "models\tessdata_fast" (
        xcopy /E /I /Y /Q "models\tessdata_fast" "%FAST_DST%" >nul
    )
    if exist "models\tessdata_best" (
        xcopy /E /I /Y /Q "models\tessdata_best" "%BEST_DST%" >nul
    )
) else (
    echo [OK] OCR language models already in place.
)

REM Download any language file still missing (e.g. if the models
REM folder was removed from this archive for some reason).
for %%L in (rus kir osd) do (
    if not exist "%FAST_DST%\%%L.traineddata" (
        echo Downloading model %%L ^(fast^)...
        powershell -NoProfile -Command "try { Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/main/%%L.traineddata' -OutFile '%FAST_DST%\%%L.traineddata' } catch { exit 1 }"
    )
    if not exist "%BEST_DST%\%%L.traineddata" (
        echo Downloading model %%L ^(best^)...
        powershell -NoProfile -Command "try { Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/tesseract-ocr/tessdata_best/main/%%L.traineddata' -OutFile '%BEST_DST%\%%L.traineddata' } catch { exit 1 }"
    )
)
echo [OK] OCR language models are ready.
echo.

REM ---------------------------------------------------------------
REM 6. Start the application
REM ---------------------------------------------------------------
echo ================================================================
echo   Starting the application...
echo   Open in your browser: http://127.0.0.1:8765
echo   To stop the server, close this window or press Ctrl+C
echo ================================================================
echo.

"%PYEXE%" -m gns_app.launcher

echo.
echo Server stopped.
pause
