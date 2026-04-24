$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = Join-Path $root "usdrate.py"
$distPath = Join-Path $root "dist_light"
$workPath = Join-Path $root "build_light"

if (-not (Test-Path $scriptPath)) {
    throw "Script not found: $scriptPath"
}

py -3 -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name usdrate `
    --distpath $distPath `
    --workpath $workPath `
    --specpath $root `
    $scriptPath

if ($LASTEXITCODE -ne 0) {
    throw "Build failed."
}

Write-Output "Build complete."
Write-Output "EXE: $(Join-Path $distPath 'usdrate.exe')"
