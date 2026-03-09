# Setup Kit (Python + Rainmeter)

This folder builds a distributable kit for `dexcom.py` and installs a Rainmeter skin.

## Files
- `build-kit.ps1`: Builds `DexcomReader.exe` with PyInstaller and creates `out/DexcomRainmeterKit.zip`.
- `install-kit.ps1`: Installs the skin and saves Dexcom credentials as user environment variables.
- `../Dexcom.ini`: Rainmeter skin template source; copied into the kit during build and placeholders are replaced during install.

## Build
From repository root:

```powershell
cd .\setup-kit
.\build-kit.ps1
```

By default, build artifacts are written to:
`setup-kit\\out`

## Install (from generated kit folder)

```powershell
cd .\out\DexcomRainmeterKit
.\install-kit.ps1
```

Alternative (from `setup-kit` after build):

```powershell
cd .\setup-kit
.\install-kit.ps1
```

`install-kit.ps1` prompts for username and password only when `DEXCOM_USER` or `DEXCOM_PASS` are missing.
`DEXCOM_REGION` is always set to `us`.
`DexcomReader.exe` is copied to `%LOCALAPPDATA%\\DexcomRainmeterKit\\app` and Rainmeter runs it from there.
Default install path uses the current user's Documents folder automatically (supports redirected OneDrive Documents).
The script also sends Rainmeter `!ActivateConfig` and `!RefreshApp`, so the skin is loaded automatically.
If Rainmeter is not running, the script launches it first.

If Rainmeter is not installed in `Program Files`, load the skin manually in Rainmeter.
