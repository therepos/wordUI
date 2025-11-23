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
- Copies wordUI.dotm into Word's Startup folder:
  %APPDATA%\Microsoft\Word\Startup
- No C:\Apps, no Trusted Locations; Word already trusts Startup.
#>

# ===== CONFIG =====
$StartupDir     = Join-Path $env:APPDATA 'Microsoft\Word\Startup'
$TargetDir      = $StartupDir          # keep for potential reuse
$FilesToInstall = @('wordUI.dotm')
$AddInFile      = 'wordUI.dotm'
# ==================

# -------- Helpers --------
function Get-ScriptDir {
    if ($env:SETUP_DIR)                 { return $env:SETUP_DIR }
    if ($SourceDirOverride)             { return $SourceDirOverride }
    if ($PSScriptRoot)                  { return $PSScriptRoot }
    if ($PSCommandPath)                 { return (Split-Path -Parent $PSCommandPath) }
    if ($MyInvocation.MyCommand.Path)   { return (Split-Path -Parent $MyInvocation.MyCommand.Path) }
    return (Get-Location).Path
}
$SourceDir = (Get-ScriptDir).TrimEnd('\') + '\'

function Ensure-Dir($p){
    if (-not (Test-Path $p)) {
        New-Item -ItemType Directory -Path $p -Force | Out-Null
    }
}

function Status($msg,[scriptblock]$act,[switch]$Fatal){
    $w = 40
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

# Detection (file must exist in Startup)
function Detect-Installed {
    $installed = @()

    $startupAddin = Join-Path $StartupDir $AddInFile
    if (Test-Path $startupAddin) { $installed += $startupAddin }

    $installed | Select-Object -Unique
}

# Actions
function Install-Addin {
    Status "Installing add-in to Word Startup" -Fatal {
        Ensure-Dir $StartupDir

        foreach($f in $FilesToInstall){
            $src = Join-Path $SourceDir $f
            if (-not (Test-Path $src)) {
                throw "Missing file '$f' in $SourceDir"
            }
            $dstStart  = Join-Path $StartupDir $f
            Copy-Item $src $dstStart -Force
        }
    }
}

function Uninstall-Addin {
    Status "Removing add-in from Word Startup" {
        $p = Join-Path $StartupDir $AddInFile
        if (Test-Path $p) {
            Remove-Item $p -Force -ErrorAction SilentlyContinue
        }
    }
}

# Main
$paths = Detect-Installed
if (-not $paths -or $paths.Count -eq 0) {
    $ans = Read-Host "wordUI is NOT installed. Install now? (Y/N)"
    if ($ans -match '^[Yy]') { Install-Addin }
} else {
    $ans = Read-Host ("wordUI is installed at: " + ($paths -join ', ') + ". Uninstall it? (Y/N)")
    if ($ans -match '^[Yy]') { Uninstall-Addin }
}

Write-Host ""
Write-Host "All done. You can close this window now." -ForegroundColor Yellow

exit 0
