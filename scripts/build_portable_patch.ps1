param(
    [Parameter(Mandatory = $true)]
    [string]$ReleaseDir,
    [string]$OutputRoot = ""
)

$ErrorActionPreference = "Stop"
$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$ReleaseDir = (Resolve-Path $ReleaseDir).Path
if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $PublicDocuments = [Environment]::GetFolderPath(
        [Environment+SpecialFolder]::CommonDocuments
    )
    $OutputRoot = Join-Path $PublicDocuments (
        "GNS-Releases\GNS-Update-" + (Get-Date -Format "yyyyMMdd-HHmmss")
    )
}
$OutputRoot = [IO.Path]::GetFullPath($OutputRoot)
if (Test-Path -LiteralPath $OutputRoot) {
    throw "Update directory already exists: $OutputRoot"
}

foreach ($Required in @(
    (Join-Path $ReleaseDir "GNS-Portable.exe"),
    (Join-Path $ReleaseDir "GNS-Worker.exe"),
    (Join-Path $ReleaseDir "_internal\gns_app"),
    (Join-Path $ProjectRoot "packaging\UPDATE_EXISTING.bat")
)) {
    if (-not (Test-Path -LiteralPath $Required)) {
        throw "Required update file was not found: $Required"
    }
}

$PatchDir = Join-Path $OutputRoot "GNS-Update"
$FilesDir = Join-Path $PatchDir "files"
New-Item -ItemType Directory -Force -Path `
    $FilesDir, (Join-Path $FilesDir "_internal") | Out-Null

Copy-Item -LiteralPath (Join-Path $ReleaseDir "GNS-Portable.exe") -Destination $FilesDir
Copy-Item -LiteralPath (Join-Path $ReleaseDir "GNS-Worker.exe") -Destination $FilesDir
Copy-Item -LiteralPath (Join-Path $ReleaseDir "_internal\gns_app") `
    -Destination (Join-Path $FilesDir "_internal\gns_app") -Recurse
foreach ($Optional in @("START.bat", "README.txt")) {
    $Source = Join-Path $ReleaseDir $Optional
    if (Test-Path -LiteralPath $Source -PathType Leaf) {
        Copy-Item -LiteralPath $Source -Destination $FilesDir
    }
}
Copy-Item -LiteralPath (Join-Path $ProjectRoot "packaging\UPDATE_EXISTING.bat") `
    -Destination (Join-Path $PatchDir "UPDATE.bat")

@"
UPDATE FOR AN EXISTING GNS-PORTABLE FOLDER

1. Extract the GNS-Update folder inside the existing GNS-Portable folder.
2. Run GNS-Update\UPDATE.bat.
3. If the application is open, close it when requested.

The runtime folder, database, history, scans, responses and Inbox folder are
not replaced. Application files and a pre-update database copy are saved to
.update-backup first.
"@ | Set-Content -LiteralPath (Join-Path $PatchDir "HOW-TO-UPDATE.txt") -Encoding ASCII

$Manifest = Get-ChildItem -LiteralPath $PatchDir -Recurse -File |
    Sort-Object FullName |
    ForEach-Object {
        [ordered]@{
            path = $_.FullName.Substring($PatchDir.Length + 1).Replace('\', '/')
            bytes = $_.Length
            sha256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    }
[ordered]@{
    build = "GNS-Update"
    compatible_with = "GNS-Portable-20260907-140229"
    created_at = (Get-Date).ToString("o")
    files = @($Manifest)
} | ConvertTo-Json -Depth 4 | Set-Content `
    -LiteralPath (Join-Path $PatchDir "update-manifest.json") -Encoding UTF8

$ArchivePath = Join-Path $OutputRoot "GNS-Update.zip"
Compress-Archive -LiteralPath $PatchDir -DestinationPath $ArchivePath -CompressionLevel Optimal
Write-Output $PatchDir
Write-Output $ArchivePath
