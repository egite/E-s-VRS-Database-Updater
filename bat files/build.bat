@echo off
setlocal enabledelayedexpansion

REM This bat lives in "<project>\bat files\" — step up one level so PyInstaller
REM runs against the project root (where run.py, Rules.csv, the icons, etc. live).
for %%I in ("%~dp0..") do set "PROJECT_DIR=%%~fI\"
cd /d "%PROJECT_DIR%"

echo ============================================================
echo  E's VRS Database Updater - Build Script
echo ============================================================
echo.

REM --- Install PyInstaller if needed ---
echo [1/5] Checking dependencies...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo ERROR: Failed to install PyInstaller.
        pause
        exit /b 1
    )
)

REM --- Common PyInstaller options ---
set "ENTRY=run.py"
set "ICON=database.ico"
set "NAME=E's VRS Database Updater"
set "COMMON_OPTS=--onefile --windowed --noconfirm --clean"
set "DATA_OPTS=--add-data "Rules.csv;." --add-data "Sils.csv;." --add-data "database.ico;." --add-data "banner_app_icon.png;." --add-data "banner_arrow.png;." --add-data "banner_sources.png;." --add-data "banner_vrs_icon.png;.""

REM --- Build 64-bit ---
echo.
echo [2/5] Building 64-bit executable...
echo ============================================================
python -m PyInstaller %COMMON_OPTS% ^
    --name "VRS_Database_Updater_x64" ^
    --icon "%ICON%" ^
    --add-data "Rules.csv;." ^
    --add-data "Sils.csv;." ^
    --add-data "database.ico;." ^
    --add-data "banner_app_icon.png;." ^
    --add-data "banner_arrow.png;." ^
    --add-data "banner_sources.png;." ^
    --add-data "banner_vrs_icon.png;." ^
    "%ENTRY%"

if errorlevel 1 (
    echo ERROR: 64-bit build failed.
    pause
    exit /b 1
)
echo 64-bit build OK.

REM --- Check for 32-bit Python ---
echo.
echo [3/5] Building 32-bit executable...
echo ============================================================

REM Try to find 32-bit Python via py launcher
set "PY32="
for /f "tokens=*" %%i in ('py -3.12-32 -c "import sys; print(sys.executable)" 2^>nul') do set "PY32=%%i"

if not defined PY32 (
    echo No 32-bit Python found. Attempting to install Python 3.12 32-bit...
    echo Downloading Python 3.12 32-bit installer...
    powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.10/python-3.12.10.exe' -OutFile '%TEMP%\python312_32.exe'"
    if errorlevel 1 (
        echo ERROR: Failed to download 32-bit Python installer.
        echo Skipping 32-bit build. You can install Python 3.12 32-bit manually.
        goto :done
    )
    echo Installing Python 3.12 32-bit...
    "%TEMP%\python312_32.exe" /quiet TargetDir="%LOCALAPPDATA%\Programs\Python\Python312-32" Include_launcher=0 InstallAllUsers=0 AssociateFiles=0 Include_doc=0 Include_test=0
    set "PY32=%LOCALAPPDATA%\Programs\Python\Python312-32\python.exe"

    if not exist "!PY32!" (
        echo ERROR: 32-bit Python installation failed.
        echo Skipping 32-bit build.
        goto :done
    )
)

REM Install PyInstaller and requests in 32-bit Python
echo Using 32-bit Python: %PY32%
"%PY32%" -m pip install pyinstaller requests --quiet 2>nul

"%PY32%" -m PyInstaller %COMMON_OPTS% ^
    --name "VRS_Database_Updater_x86" ^
    --icon "%ICON%" ^
    --add-data "Rules.csv;." ^
    --add-data "Sils.csv;." ^
    --add-data "database.ico;." ^
    --add-data "banner_app_icon.png;." ^
    --add-data "banner_arrow.png;." ^
    --add-data "banner_sources.png;." ^
    --add-data "banner_vrs_icon.png;." ^
    "%ENTRY%"

if errorlevel 1 (
    echo ERROR: 32-bit build failed.
    goto :done
)
echo 32-bit build OK.

:done
echo.
echo [4/5] Build Summary
echo ============================================================
if exist "dist\VRS_Database_Updater_x64.exe" (
    echo   x64: dist\VRS_Database_Updater_x64.exe
    for %%F in ("dist\VRS_Database_Updater_x64.exe") do echo        Size: %%~zF bytes
) else (
    echo   x64: FAILED
)
if exist "dist\VRS_Database_Updater_x86.exe" (
    echo   x86: dist\VRS_Database_Updater_x86.exe
    for %%F in ("dist\VRS_Database_Updater_x86.exe") do echo        Size: %%~zF bytes
) else (
    echo   x86: NOT BUILT
)
echo.
echo Output folder: %PROJECT_DIR%dist\
echo ============================================================

