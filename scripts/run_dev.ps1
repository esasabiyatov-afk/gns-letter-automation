$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $projectRoot

$pythonPath = Join-Path $projectRoot ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $pythonPath)) {
    throw "Virtual environment not found. Run .\scripts\setup.ps1 first."
}

Write-Host "Application: http://127.0.0.1:8765"
Write-Host "Press Ctrl+C to stop."
& $pythonPath -m gns_app.main
