param(
    [string]$RainmeterSkinRoot = $(Join-Path ([Environment]::GetFolderPath("MyDocuments")) "Rainmeter\Skins\illustro"),
    [string]$SkinName = "DexcomWidget",
    [string]$AppInstallRoot = $(Join-Path $env:LOCALAPPDATA "DexcomRainmeterKit\app")
)

$ErrorActionPreference = "Stop"

function Get-ExistingEnvValue {
    param([string]$Name)

    $value = [Environment]::GetEnvironmentVariable($Name, "Process")
    if (-not $value) {
        $value = [Environment]::GetEnvironmentVariable($Name, "User")
    }
    if (-not $value) {
        $value = [Environment]::GetEnvironmentVariable($Name, "Machine")
    }

    return $value
}

function ConvertTo-PlainText {
    param([securestring]$SecureValue)

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureValue)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        if ($bstr -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}

function Refresh-Rainmeter {
    param(
        [string]$ConfigPath,
        [string]$IniFile = "Dexcom.ini"
    )

    $exeCandidates = @(
        "$env:ProgramFiles\Rainmeter\Rainmeter.exe",
        "$env:ProgramFiles(x86)\Rainmeter\Rainmeter.exe"
    )
    $rainmeterExe = $exeCandidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1

    if (-not $rainmeterExe) {
        Write-Host "Rainmeter.exe not found in Program Files. Refresh Rainmeter manually if needed."
        return
    }

    $isRunning = @(Get-Process -Name "Rainmeter" -ErrorAction SilentlyContinue).Count -gt 0
    if (-not $isRunning) {
        Start-Process -FilePath $rainmeterExe | Out-Null
        Start-Sleep -Seconds 2
        Write-Host "Rainmeter was not running; launched it."
    }

    if ($ConfigPath) {
        & $rainmeterExe "!ActivateConfig" $ConfigPath $IniFile
        Write-Host "Rainmeter activate command sent for: $ConfigPath\\$IniFile"
    }

    & $rainmeterExe "!RefreshApp"
    Write-Host "Rainmeter refresh command sent."
}

function Get-RainmeterConfigPath {
    param([string]$SkinDir)

    $skinsRoot = Join-Path ([Environment]::GetFolderPath("MyDocuments")) "Rainmeter\Skins"
    if ($SkinDir.StartsWith($skinsRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        $relative = $SkinDir.Substring($skinsRoot.Length).TrimStart("\\")
        if ($relative) {
            return $relative
        }
    }

    return $SkinName
}

$packageRoot = $PSScriptRoot
$repoRoot = (Resolve-Path (Join-Path $packageRoot "..")).Path

$exeCandidates = @(
    (Join-Path $packageRoot "app\DexcomReader.exe"),
    (Join-Path $packageRoot "out\DexcomRainmeterKit\app\DexcomReader.exe"),
    (Join-Path $packageRoot "out\dist\DexcomReader.exe"),
    (Join-Path $env:LOCALAPPDATA "DexcomRainmeterKit\build\DexcomRainmeterKit\app\DexcomReader.exe"),
    (Join-Path $env:LOCALAPPDATA "DexcomRainmeterKit\build\dist\DexcomReader.exe")
)
$sourceExePath = $exeCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $sourceExePath) {
    throw "DexcomReader.exe not found. Run .\build-kit.ps1 first, then run .\out\DexcomRainmeterKit\install-kit.ps1 (or rerun this script from setup-kit after build)."
}

New-Item -Path $AppInstallRoot -ItemType Directory -Force | Out-Null
$installedExePath = Join-Path $AppInstallRoot "DexcomReader.exe"
Copy-Item -Path $sourceExePath -Destination $installedExePath -Force

$templateCandidates = @(
    (Join-Path $packageRoot "skin\Dexcom.ini"),
    (Join-Path $repoRoot "Dexcom.ini")
)
$templateIniPath = $templateCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $templateIniPath) {
    throw "Dexcom.ini template not found. Expected one of: skin\Dexcom.ini or ..\Dexcom.ini"
}

$targetSkinDir = Join-Path $RainmeterSkinRoot $SkinName
New-Item -Path $targetSkinDir -ItemType Directory -Force | Out-Null

$iniContent = Get-Content -Path $templateIniPath -Raw
$programValue = $installedExePath -replace "\\", "/"
$iniContent = $iniContent.Replace("__CMD_PROGRAM__", $programValue)
$iniContent = $iniContent.Replace("__CMD_PARAM__", "")

$targetIniPath = Join-Path $targetSkinDir "Dexcom.ini"
Set-Content -Path $targetIniPath -Value $iniContent -Encoding UTF8

Write-Host "Installed skin to: $targetSkinDir"
Write-Host "Installed app to: $installedExePath"
$configPath = Get-RainmeterConfigPath -SkinDir $targetSkinDir

$dexcomUser = Get-ExistingEnvValue -Name "DEXCOM_USER"
$userWasProvided = $false
if (-not $dexcomUser) {
    $dexcomUser = Read-Host -Prompt "Enter Dexcom username"
    if (-not $dexcomUser) {
        throw "Dexcom username is required."
    }
    $userWasProvided = $true
}

$existingPass = Get-ExistingEnvValue -Name "DEXCOM_PASS"
$passwordWasProvided = $false
if (-not $existingPass) {
    $securePass = Read-Host -Prompt "Enter Dexcom password" -AsSecureString
    $existingPass = ConvertTo-PlainText -SecureValue $securePass
    if (-not $existingPass) {
        throw "Dexcom password is required."
    }
    $passwordWasProvided = $true
}

[Environment]::SetEnvironmentVariable("DEXCOM_REGION", "us", "User")

if ($userWasProvided) {
    [Environment]::SetEnvironmentVariable("DEXCOM_USER", $dexcomUser, "User")
}

if ($passwordWasProvided) {
    [Environment]::SetEnvironmentVariable("DEXCOM_PASS", $existingPass, "User")
}

if ($userWasProvided -and $passwordWasProvided) {
    Write-Host "Saved DEXCOM_USER, DEXCOM_PASS, and DEXCOM_REGION=us at user scope."
} elseif ($userWasProvided -and -not $passwordWasProvided) {
    Write-Host "Saved DEXCOM_USER and DEXCOM_REGION=us at user scope. Existing DEXCOM_PASS was preserved."
} elseif (-not $userWasProvided -and $passwordWasProvided) {
    Write-Host "Saved DEXCOM_PASS and DEXCOM_REGION=us at user scope. Existing DEXCOM_USER was preserved."
} else {
    Write-Host "Saved DEXCOM_REGION=us at user scope. Existing DEXCOM_USER and DEXCOM_PASS were preserved."
}

Refresh-Rainmeter -ConfigPath $configPath -IniFile "Dexcom.ini"
Write-Host "Done. Skin should be active: '$configPath\\Dexcom.ini'."
