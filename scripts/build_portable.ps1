param(
    [switch]$SkipTests,
    [string]$OutputRoot = ""
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$PythonExe = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$SpecFile = Join-Path $ProjectRoot "packaging\gns_portable.spec"
if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $PublicDocuments = [Environment]::GetFolderPath(
        [Environment+SpecialFolder]::CommonDocuments
    )
    if ([string]::IsNullOrWhiteSpace($PublicDocuments)) {
        throw "Не удалось определить отдельную папку для portable-сборки. Укажите -OutputRoot."
    }
    $OutputRoot = Join-Path $PublicDocuments (
        "GNS-Releases\GNS-Portable-" + (Get-Date -Format "yyyyMMdd-HHmmss")
    )
}
$ReleaseRoot = [IO.Path]::GetFullPath($OutputRoot)
if (Test-Path -LiteralPath $ReleaseRoot) {
    throw "Папка отдельной сборки уже существует: $ReleaseRoot"
}
$DistRoot = Join-Path $ReleaseRoot "dist"
$BuildRoot = Join-Path $ReleaseRoot "pyinstaller-work"
$ReleaseDir = Join-Path $DistRoot "GNS-Portable"
$ArchivePath = Join-Path $ReleaseRoot "GNS-Portable.zip"

foreach ($Required in @(
    $PythonExe,
    $SpecFile,
    (Join-Path $ProjectRoot "models\tessdata_fast\rus.traineddata"),
    (Join-Path $ProjectRoot "models\tessdata_fast\kir.traineddata"),
    (Join-Path $ProjectRoot "models\tessdata_best\rus.traineddata"),
    (Join-Path $ProjectRoot "models\tessdata_best\kir.traineddata"),
    (Join-Path $ProjectRoot "УГНС\шаблон ответа одиночный.docx"),
    (Join-Path $ProjectRoot "УГНС\шаблон ответа много.docx"),
    (Join-Path $ProjectRoot "packaging\assets\gns-document-seal.ico"),
    (Join-Path $ProjectRoot "packaging\assets\gns-folder-mail.ico"),
    (Join-Path $ProjectRoot "packaging\assets\gns-shield-document.ico")
)) {
    if (-not (Test-Path -LiteralPath $Required -PathType Leaf)) {
        throw "Обязательный файл сборки не найден: $Required"
    }
}

# Windows 8.1 is the supported target for this portable edition.  Do not let a
# later local virtual environment silently produce an incompatible build.
$PythonTarget = & $PythonExe -c "import struct, sys; print('%s|%s.%s|%s' % (sys.implementation.name, sys.version_info.major, sys.version_info.minor, struct.calcsize('P') * 8))"
if ($LASTEXITCODE -ne 0 -or $PythonTarget.Trim() -ne "cpython|3.12|64") {
    throw "GNS-Portable для Windows 8.1 нужно собирать CPython 3.12 x64; найдено: $PythonTarget"
}

if (-not $SkipTests) {
    & $PythonExe -m pytest
    if ($LASTEXITCODE -ne 0) { throw "pytest failed" }
}

New-Item -ItemType Directory -Force -Path $DistRoot, $BuildRoot | Out-Null

& $PythonExe -m PyInstaller `
    --noconfirm `
    --clean `
    --distpath $DistRoot `
    --workpath $BuildRoot `
    $SpecFile
if ($LASTEXITCODE -ne 0) { throw "PyInstaller build failed" }

foreach ($BuiltFile in @("GNS-Portable.exe", "GNS-Worker.exe")) {
    if (-not (Test-Path -LiteralPath (Join-Path $ReleaseDir $BuiltFile) -PathType Leaf)) {
        throw "Служебный файл сборки не найден: $BuiltFile"
    }
}

Copy-Item -LiteralPath (Join-Path $ProjectRoot "packaging\START_PORTABLE.bat") `
    -Destination (Join-Path $ReleaseDir "START.bat")
Copy-Item -LiteralPath (Join-Path $ProjectRoot "packaging\PORTABLE_README_RU.txt") `
    -Destination (Join-Path $ReleaseDir "README.txt")
Copy-Item -LiteralPath (Join-Path $ProjectRoot "packaging\assets") `
    -Destination (Join-Path $ReleaseDir "Иконки") -Recurse
New-Item -ItemType Directory -Force -Path `
    (Join-Path $ReleaseDir "runtime"), `
    (Join-Path $ReleaseDir "Входящие") | Out-Null

$Manifest = Get-ChildItem -LiteralPath $ReleaseDir -Recurse -File |
    Where-Object { $_.Name -ne "release-manifest.json" } |
    Sort-Object FullName |
    ForEach-Object {
        [ordered]@{
            path = $_.FullName.Substring($ReleaseDir.Length + 1).Replace('\', '/')
            bytes = $_.Length
            sha256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    }
[ordered]@{
    build = "GNS-Portable"
    created_at = (Get-Date).ToString("o")
    files = @($Manifest)
} | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $ReleaseDir "release-manifest.json") -Encoding UTF8

Compress-Archive -LiteralPath $ReleaseDir -DestinationPath $ArchivePath -CompressionLevel Optimal

Write-Output $ReleaseDir
Write-Output $ArchivePath
