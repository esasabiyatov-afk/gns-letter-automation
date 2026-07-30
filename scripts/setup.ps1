$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $projectRoot

if (-not (Test-Path -LiteralPath ".venv\Scripts\python.exe")) {
    python -m venv .venv
}

& ".venv\Scripts\python.exe" -m pip install --upgrade pip
& ".venv\Scripts\python.exe" -m pip install -e ".[dev]"

$pythonPath = Join-Path $projectRoot ".venv\Scripts\python.exe"
& $pythonPath -c "import tesserocr" 2>$null
if ($LASTEXITCODE -ne 0) {
    $wheelUrl = "https://github.com/simonflueckiger/tesserocr-windows_build/releases/download/tesserocr-v2.10.0-tesseract-5.5.2/tesserocr-2.10.0-cp312-cp312-win_amd64.whl"
    & $pythonPath -m pip install $wheelUrl
}

$modelRoot = Join-Path $env:LOCALAPPDATA "GNSLetterAutomation\models"
$fastRoot = Join-Path $modelRoot "tessdata_fast"
$bestRoot = Join-Path $modelRoot "tessdata_best"
New-Item -ItemType Directory -Force -Path $fastRoot | Out-Null
New-Item -ItemType Directory -Force -Path $bestRoot | Out-Null

function Install-OcrModel {
    param(
        [string]$Repository,
        [string]$Language,
        [string]$Destination
    )
    $target = Join-Path $Destination "$Language.traineddata"
    if (-not (Test-Path -LiteralPath $target)) {
        $url = "https://raw.githubusercontent.com/tesseract-ocr/$Repository/main/$Language.traineddata"
        Invoke-WebRequest -Uri $url -OutFile $target
    }
}

foreach ($language in @("rus", "kir", "osd")) {
    Install-OcrModel -Repository "tessdata_fast" -Language $language -Destination $fastRoot
    Install-OcrModel -Repository "tessdata_best" -Language $language -Destination $bestRoot
}

Write-Host "Setup completed. Run: .\scripts\run_dev.ps1"
