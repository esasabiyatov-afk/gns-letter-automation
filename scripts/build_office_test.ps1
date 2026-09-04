param(
    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$PythonExe = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$SpecFile = Join-Path $ProjectRoot "packaging\gns_office_test.spec"
$DistRoot = Join-Path $ProjectRoot "dist"
$BuildRoot = Join-Path $ProjectRoot "build\pyinstaller-office-test"
$ReleaseDir = Join-Path $DistRoot "GNS-Test-Win8.1"

if (-not (Test-Path -LiteralPath $PythonExe -PathType Leaf)) {
    throw "Python virtual environment was not found: $PythonExe"
}

if (-not $SkipTests) {
    & $PythonExe -m pytest
    if ($LASTEXITCODE -ne 0) {
        throw "pytest failed"
    }
}

New-Item -ItemType Directory -Force -Path `
    $DistRoot, `
    (Split-Path $BuildRoot -Parent) | Out-Null

foreach ($Target in @($BuildRoot, $ReleaseDir)) {
    $ResolvedParent = (Resolve-Path (Split-Path $Target -Parent)).Path
    if (-not $ResolvedParent.StartsWith($ProjectRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Unsafe build target: $Target"
    }
    if (Test-Path -LiteralPath $Target) {
        Remove-Item -LiteralPath $Target -Recurse -Force
    }
}

& $PythonExe -m PyInstaller `
    --noconfirm `
    --clean `
    --distpath $DistRoot `
    --workpath $BuildRoot `
    $SpecFile
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller build failed"
}

Copy-Item -LiteralPath `
    (Join-Path $ProjectRoot "packaging\OFFICE_TEST_README_RU.txt") `
    -Destination (Join-Path $ReleaseDir "OFFICE_TEST_README_RU.txt")
Copy-Item -LiteralPath (Join-Path $ProjectRoot "packaging\assets") `
    -Destination (Join-Path $ReleaseDir "Иконки") -Recurse
New-Item -ItemType Directory -Force -Path (Join-Path $ReleaseDir "runtime") | Out-Null

Write-Output $ReleaseDir
