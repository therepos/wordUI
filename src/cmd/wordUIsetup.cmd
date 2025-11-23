@echo off
setlocal EnableExtensions

:: Sanity: required files must be beside this cmd
if not exist "%~dp0wordUI.dotm" (
  echo Missing wordUI.dotm next to wordUIsetup.cmd
  pause
  exit /b 1
)

:: Extract embedded PowerShell payload
set "PAYTAG=::PAYLOAD"
for /f "delims=:" %%A in ('findstr /n /c:"%PAYTAG%" "%~f0"') do set /a LN=%%A+1
set "TMPPS=%TEMP%\wordUI_setup_%RANDOM%.ps1"
more +%LN% "%~f0" > "%TMPPS%"

:: Pass our folder via ENV and PS param; run STA for Office COM if needed
set "SETUP_DIR=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Sta -File "%TMPPS%" -SourceDirOverride "%~dp0"
set "rc=%ERRORLEVEL%"
del "%TMPPS%" >nul 2>&1
if not "%rc%"=="0" (
  echo.
  echo Installer reported an error (code %rc%). See messages above.
  pause
)
endlocal
exit /b %rc%

::PAYLOAD
param([string]$SourceDirOverride)

$ErrorActionPreference = 'Stop'
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
$Host.UI.RawUI.WindowTitle = 'wordUI setup'

<# 
Word Add-in Installer/Uninstaller (single file)
- Copies to C:\Apps (creates if missing)
- Adds Word Trusted Location (HKCU for 16.0/15.0/14.0 if present)
- Copies wordUI.dotm to Word STARTUP folder so it autoloads
#>

# ===== CONFIG =====
$TargetDir      = 'C:\Apps'
$FilesToInstall = @('wordUI.dotm')
$AddInFile      = 'wordUI.dotm'
$TrustedDesc    = 'Installer-Managed: Apps Word Add-in'
$StartupDir     = Join-Path $env:APPDATA 'Microsoft\Word\Startup'
# ==================

# -------- Helpers --------
function Get-ScriptDir {
    if ($env:SETUP_DIR)           { return $env:SETUP_DIR }
    if ($SourceDirOverride)       { return $SourceDirOverride }
    if ($PSScriptRoot)            { return $PSScriptRoot }
    if ($PSCommandPath)           { return (Split-Path -Parent $PSCommandPath) }
    if ($MyInvocation.MyCommand.Path) { return (Split-Path -Parent $MyInvocation.MyCommand.Path) }
    return (Get-Location).Path
}
$SourceDir = (Get-ScriptDir).TrimEnd('\') + '\'

function Ensure-Dir($p){
    if (-not (Test-Path $p)) {
        New-Item -ItemType Directory -Path $p -Force | Out-Null
    }
}

function Get-OfficeVersions {
    @('16.0','15.0','14.0') | Where-Object {
        Test-Path "HKCU:\Software\Microsoft\Office\$_"
    }
}

function Status($msg,[scriptblock]$act,[switch]$Fatal){
    $w = 34
    Write-Host ($msg.PadRight($w)) -NoNewline
    try{
        & $act | Out-Null
        Write-Host "Done" -ForegroundColor Green
    }
    catch{
        Write-Host "Failed" -ForegroundColor Red
        Write-Host ("  " + $_.Exception.Message) -ForegroundColor DarkRed
        if($Fatal){ throw }
    }
}

# Trusted Locations (Word)
function Add-TrustedLocation($path,$desc){
    $ok = $false
    foreach($ver in Get-OfficeVersions){
        $base = "HKCU:\Software\Microsoft\Office\$ver\Word\Security\Trusted Locations"
        if (-not (Test-Path $base)) { continue }
        $n = 1
        while (Test-Path "$base\Location$n") { $n++ }
        $k = "$base\Location$n"
        New-Item -Path $k -Force | Out-Null
        New-ItemProperty -Path $k -Name Path -Value ($path.TrimEnd('\') + '\') -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $k -Name AllowSubFolders -Value 1 -PropertyType DWord -Force | Out-Null
        New-ItemProperty -Path $k -Name Description -Value $desc -PropertyType String -Force | Out-Null
        $ok = $true
    }
    return $ok
}

function Remove-TrustedLocation($path,$desc){
    $removed = $false
    foreach($ver in Get-OfficeVersions){
        $base = "HKCU:\Software\Microsoft\Office\$ver\Word\Security\Trusted Locations"
        if (-not (Test-Path $base)) { continue }
        Get-ChildItem $base | ForEach-Object{
            $p = (Get-ItemProperty $_.PsPath -ErrorAction SilentlyContinue)
            if ($p.Path -and ($p.Path.TrimEnd('\')+'\') -ieq ($path.TrimEnd('\')+'\') -and ($p.Description -like "$desc*")) {
                Remove-Item $_.PsPath -Recurse -Force -ErrorAction SilentlyContinue
                $removed = $true
            }
        }
    }
    return $removed
}

# Detection (file must exist)
function Detect-Installed {
    $installed = @()

    $startupAddin = Join-Path $StartupDir $AddInFile
    if (Test-Path $startupAddin) { $installed += $startupAddin }

    $appsAddin = Join-Path $TargetDir $AddInFile
    if (Test-Path $appsAddin) { $installed += $appsAddin }

    $installed | Select-Object -Unique
}

# Actions
function Install-Addin {
    Status "Installing files" -Fatal {
        Ensure-Dir $TargetDir
        Ensure-Dir $StartupDir

        foreach($f in $FilesToInstall){
            $src = Join-Path $SourceDir $f
            if (-not (Test-Path $src)) {
                throw "Missing file '$f' in $SourceDir"
            }
            $dstApps   = Join-Path $TargetDir  $f
            $dstStart  = Join-Path $StartupDir $f

            Copy-Item $src $dstApps  -Force
            Copy-Item $src $dstStart -Force
        }
    }

    Status "Adding Trusted Location" {
        if (-not (Add-TrustedLocation -path $TargetDir -desc $TrustedDesc)) {
            throw "Could not add trusted location"
        }
    }
}

function Uninstall-Addin($paths){
    Status "Removing Startup add-in" {
        $p = Join-Path $StartupDir $AddInFile
        if (Test-Path $p) {
            Remove-Item $p -Force -ErrorAction SilentlyContinue
        }
    }

    Status "Removing C:\Apps files" {
        foreach($f in $FilesToInstall){
            $dst = Join-Path $TargetDir $f
            if (Test-Path $dst) {
                Remove-Item $dst -Force -ErrorAction SilentlyContinue
            }
        }
    }

    Status "Removing Trusted Location" {
        Remove-TrustedLocation -path $TargetDir -desc $TrustedDesc | Out-Null
    }
}

# Main
$paths = Detect-Installed
if (-not $paths -or $paths.Count -eq 0) {
    $ans = Read-Host "wordUI is NOT installed. Install now? (Y/N)"
    if ($ans -match '^[Yy]') { Install-Addin }
} else {
    $ans = Read-Host ("wordUI is installed at: " + ($paths -join ', ') + ". Uninstall it? (Y/N)")
    if ($ans -match '^[Yy]') { Uninstall-Addin $paths }
}

Write-Host ""
Write-Host "All done. You can close this window now." -ForegroundColor Yellow

exit 0
