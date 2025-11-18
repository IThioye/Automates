@echo off
CHCP 65001 > NUL
setlocal enabledelayedexpansion

REM == Set the base directory to the location of this script ==
REM This allows the script to be run from any location without issues
set "BASEDIR=%~dp0"
cd /d "%BASEDIR%"  

python "utils\sync_version.py"
echo ------------------------------------------------
REM === Load variables from .env ===
for /f "usebackq tokens=1,* delims==" %%A in ("config.env") do (
    set "%%A=%%B"
)
REM === You should get the following variables: VERSION, SERVER_PATH, PROJECT, ICON, MAIN_APP_FILE, SERVER_VERSION, BUILD_MODE ===
echo Loaded variables:
echo VERSION: %VERSION%
echo SERVER_PATH: %SERVER_PATH%
echo PROJECT: %PROJECT%
echo ICON: %ICON%
echo MAIN_APP_FILE: %MAIN_APP_FILE%
echo SERVER_VERSION: %SERVER_VERSION%
echo ------------------------------------------------
REM === Your base variables ===

set "ASSETS_FOLDER=assets"
set "DIST_PATH=versions\%VERSION%"
if not exist "%DIST_PATH%" mkdir "%DIST_PATH%"
set "WORK_PATH=build"
if not exist "%WORK_PATH%" mkdir "%WORK_PATH%"
set "APP_NAME=Automate %PROJECT%"
set "ADD_DATA="
set "OPTIONAL_FLAGS="

REM === Check if ASSETS_FOLDER exists ===
if exist "%ASSETS_FOLDER%" (
    set "OPTIONAL_FLAGS=!OPTIONAL_FLAGS! --add-data ""%ASSETS_FOLDER%;assets"""
) else (
    echo Asset folder '%ASSETS_FOLDER%' not found. Skipping asset bundle.
)

REM === Check if ICON is not empty ===
if not "%ICON%"=="" if /i not "%ICON%"=="None" (
    set "OPTIONAL_FLAGS=!OPTIONAL_FLAGS! --icon ""%ICON%"""
)

echo Detected version: !VERSION!

REM === Build the app ===

echo Building %APP_NAME% version !VERSION!...
pyinstaller "%MAIN_APP_FILE%" --name "%APP_NAME%" !OPTIONAL_FLAGS! --noconsole %BUILD_MODE% --distpath "%DIST_PATH%" --workpath "%WORK_PATH%"

@REM rem Check for errors in the build process
if !errorlevel! neq 0 ( 
    echo PyInstaller build failed!
    pause
    exit /b !errorlevel!
)

echo Build completed successfully!

echo ------------------------------------------------

echo Copying to server release folder...
if /i "!BUILD_MODE!"=="release" (
    set "SOURCE_DIR=versions\!VERSION!\%APP_NAME%"
) else (
    set "SOURCE_DIR=versions\!VERSION!"
)
set "TARGET_RELEASE_DIR=%SERVER_PATH%\%PROJECT%\Application\versions\!VERSION!"

echo Source directory: !SOURCE_DIR!
echo Target directory: !TARGET_RELEASE_DIR!

rem Create the versioned directory on the server
if not exist "!TARGET_RELEASE_DIR!" mkdir "!TARGET_RELEASE_DIR!"

if !errorlevel! neq 0 (
    echo Folder creation failed!
    pause
    exit /b !errorlevel!
)

rem Use robocopy for the copy
robocopy "!SOURCE_DIR!" "!TARGET_RELEASE_DIR!" /E /R:3 /W:10


echo Copy completed successfully!

echo ------------------------------------------------

rem Update the current version file on the server
echo Updating current_version.txt on server...
echo !VERSION! > "%SERVER_PATH%\%PROJECT%\Application\%SERVER_VERSION%"

rem Check for errors in updating the version file
if !errorlevel! neq 0 (
    echo Failed to update %SERVER_VERSION% !
    pause
    exit /b !errorlevel!
)

rem Update past versions file
echo Adding version !VERSION! to past_versions.txt...
echo. >> "utils\past_versions.txt"
echo !VERSION! >> "utils\past_versions.txt

rem Clean previous build and spec file
echo Cleaning up previous build folder and spec file...
rmdir /s /q "%WORK_PATH%"
del "%APP_NAME%.spec"
echo Cleaned up previous build files.
echo ------------------------------------------------

echo Version !VERSION! deployed successfully to %TARGET_RELEASE_DIR%
pause