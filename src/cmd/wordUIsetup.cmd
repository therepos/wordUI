@echo off
setlocal EnableExtensions

:: wordUI Setup — one-click install/update/uninstall
:: Downloads wordUI.dotm from GitHub and installs to Word's Startup folder.
:: Word auto-loads all .dotm files in Startup — no registry needed.
:: No admin rights needed.

set "ADDIN=wordUI.dotm"
set "ADDIN_NAME=wordUI"
set "DEST=%APPDATA%\Microsoft\Word\Startup"
set "DST=%DEST%\%ADDIN%"
set "DOWNLOAD_URL=https://raw.githubusercontent.com/therepos/wordUI/refs/heads/main/src/dotm/wordUI.dotm"

echo.
echo  ========================================
echo    wordUI - Word Add-in Setup
echo  ========================================
echo.

:: Check if already installed
if exist "%DST%" goto :existing

:: ---- FRESH INSTALL ----
:install

call :checkword
if %errorlevel%==1 exit /b 1

echo  Downloading %ADDIN% ...
call :download
if %errorlevel%==1 exit /b 1

echo.
echo  Installed. The WordUI tab will appear next time
echo  you open Word.
echo.
pause
exit /b 0

:: ---- ALREADY INSTALLED ----
:existing

echo  wordUI is already installed at:
echo    %DST%
echo.
echo  [U] Update    - download latest version
echo  [R] Uninstall - remove add-in
echo  [C] Cancel
echo.
set /p "ANS=  Choose (U/R/C): "
if /i "%ANS%"=="U" goto :update
if /i "%ANS%"=="R" goto :uninstall
echo  Cancelled.
echo.
pause
exit /b 0

:: ---- UPDATE ----
:update

call :checkword
if %errorlevel%==1 exit /b 1

echo  Downloading latest %ADDIN% ...
call :download
if %errorlevel%==1 exit /b 1

echo.
echo  Updated. Restart Word to load the new version.
echo.
pause
exit /b 0

:: ---- UNINSTALL ----
:uninstall

call :checkword
if %errorlevel%==1 exit /b 1

del /f "%DST%" >nul 2>&1

if exist "%DST%" (
    echo  ERROR: Could not remove %ADDIN%.
    echo.
    pause
    exit /b 1
)

echo.
echo  wordUI uninstalled.
echo.
pause
exit /b 0

:: ===========================================================
::  SUBROUTINES
:: ===========================================================

:checkword
tasklist /fi "imagename eq WINWORD.EXE" 2>nul | find /i "WINWORD.EXE" >nul
if %errorlevel%==0 (
    echo  Word is running. Please close it first.
    echo.
    pause
    exit /b 1
)
exit /b 0

:download
if not exist "%DEST%" mkdir "%DEST%"
powershell -NoProfile -Command "$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13; Invoke-WebRequest -Uri '%DOWNLOAD_URL%' -OutFile '%DST%'; if(!(Test-Path '%DST%')){exit 1}"
if errorlevel 1 (
    echo  Download failed. Check your internet connection.
    echo.
    pause
    exit /b 1
)
echo  Downloaded successfully.
exit /b 0