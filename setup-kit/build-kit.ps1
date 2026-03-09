param(
    [string]$PythonExe = "python",
    [string]$OutputRoot = $(Join-Path $PSScriptRoot "out")
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$outRoot = $OutputRoot
$buildRoot = Join-Path $outRoot "build"
$distRoot = Join-Path $outRoot "dist"
$kitRoot = Join-Path $outRoot "DexcomRainmeterKit"
$appOut = Join-Path $kitRoot "app"
$skinOut = Join-Path $kitRoot "skin"

$pyFile = Join-Path $repoRoot "dexcom.py"
$templateFile = Join-Path $repoRoot "Dexcom.ini"
$installScript = Join-Path $PSScriptRoot "install-kit.ps1"
$zipPath = Join-Path $outRoot "DexcomRainmeterKit.zip"

if (-not (Test-Path $pyFile)) {
    throw "Python script not found: $pyFile"
}
if (-not (Test-Path $templateFile)) {
    throw "Root Dexcom.ini not found: $templateFile"
}
if (-not (Test-Path $installScript)) {
    throw "Installer script not found: $installScript"
}

New-Item -Path $outRoot -ItemType Directory -Force | Out-Null

if (Test-Path $kitRoot) { Remove-Item $kitRoot -Recurse -Force }
if (Test-Path $buildRoot) { Remove-Item $buildRoot -Recurse -Force }
if (Test-Path $distRoot) { Remove-Item $distRoot -Recurse -Force }
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

New-Item -Path $appOut -ItemType Directory -Force | Out-Null
New-Item -Path $skinOut -ItemType Directory -Force | Out-Null

& $PythonExe -m pip install --upgrade pip
& $PythonExe -m pip install pyinstaller pydexcom

& $PythonExe -m PyInstaller `
    --onefile `
    --name DexcomReader `
    --clean `
    --distpath $distRoot `
    --workpath $buildRoot `
    --specpath $buildRoot `
    $pyFile

$builtExe = Join-Path $distRoot "DexcomReader.exe"
if (-not (Test-Path $builtExe)) {
    throw "Build failed. Missing: $builtExe"
}

Copy-Item -Path $builtExe -Destination (Join-Path $appOut "DexcomReader.exe") -Force
Copy-Item -Path $templateFile -Destination (Join-Path $skinOut "Dexcom.ini") -Force
Copy-Item -Path $installScript -Destination (Join-Path $kitRoot "install-kit.ps1") -Force

$readme = @"
Dexcom + Rainmeter Setup Kit

Contents
- app/DexcomReader.exe
- skin/Dexcom.ini
- install-kit.ps1

Install
1. Open PowerShell in this folder.
2. Run:
    .\install-kit.ps1
    If DEXCOM_USER or DEXCOM_PASS are not set, the script prompts for them.
    DEXCOM_REGION is always set to us.
    DexcomReader.exe is copied to %LOCALAPPDATA%\\DexcomRainmeterKit\\app.
3. The installer sends !ActivateConfig and !RefreshApp so the skin is loaded automatically.
4. Default install path uses the current user's Documents folder.

Build output location
- setup-kit\\out
"@

Set-Content -Path (Join-Path $kitRoot "README.txt") -Value $readme -Encoding ASCII

Compress-Archive -Path (Join-Path $kitRoot "*") -DestinationPath $zipPath

Write-Host "Kit created at: $kitRoot"
Write-Host "Zip created at: $zipPath"
