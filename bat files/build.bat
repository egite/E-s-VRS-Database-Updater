@echo off
setlocal enabledelayedexpansion

REM This bat lives in "<project>\bat files\" — step up one level so PyInstaller
REM runs against the project root (where run.py, the icons, etc. live).
for %%I in ("%~dp0..") do set "PROJECT_DIR=%%~fI\"
cd /d "%PROJECT_DIR%"

REM Capture build start time. 'o' (round-trip) format parses culture-invariantly
REM later, regardless of the machine's date/time locale.
for /f "usebackq delims=" %%s in (`powershell -NoProfile -Command "(Get-Date).ToString('o')"`) do set "BUILD_START=%%s"

echo ============================================================
echo  E's VRS Database Updater - Build Script
echo ============================================================
echo.

REM --- Install PyInstaller if needed ---
REM Go through "python -m pip" rather than pip.exe: Device Guard / WDAC policies
REM can block the pip.exe launcher shim while still allowing python.exe.
echo [1/4] Checking dependencies...
python -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    python -m pip install pyinstaller
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
set "DATA_OPTS=--add-data "database.ico;." --add-data "banner_app_icon.png;." --add-data "banner_arrow.png;." --add-data "banner_sources.png;." --add-data "banner_vrs_icon.png;.""

REM --- Build 64-bit ---
echo.
echo [2/4] Building 64-bit executable...
echo ============================================================
python -m PyInstaller %COMMON_OPTS% ^
    --name "VRS_Database_Updater_x64" ^
    --icon "%ICON%" ^
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

echo.
echo [3/4] Build Summary
echo ============================================================
if exist "dist\VRS_Database_Updater_x64.exe" (
    echo   x64: dist\VRS_Database_Updater_x64.exe
    for %%F in ("dist\VRS_Database_Updater_x64.exe") do echo        Size: %%~zF bytes
) else (
    echo   x64: FAILED
)
echo.
echo Output folder: %PROJECT_DIR%dist\
echo ============================================================


echo.
echo [4/4] Copying to J:

copy /y "%PROJECT_DIR%dist\VRS_Database_Updater_x64.exe" "J:\VRS Database Updater\VRS_Database_Updater.exe"

REM --- Build-time pop-up (finish timestamp + elapsed time) ---
powershell -NoProfile -Command "$start=[datetime]::Parse('%BUILD_START%'); $end=Get-Date; $s=$end-$start; $msg='Build completed.' + [Environment]::NewLine + [Environment]::NewLine + 'Finished: ' + $end.ToString('yyyy-MM-dd HH:mm:ss') + [Environment]::NewLine + ('Elapsed:  {0}h {1:00}m {2:00}s' -f [int]$s.TotalHours, $s.Minutes, $s.Seconds); Add-Type -AssemblyName System.Windows.Forms | Out-Null; [System.Windows.Forms.MessageBox]::Show($msg, 'VRS Database Updater - Build Complete', 'OK', 'Information') | Out-Null"

timeout /t 5