# build.ps1
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# === Base directory of the script ===
$BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $BaseDir

# === Run the Python sync script ===
python ".\utils\sync_version.py"

Write-Host "------------------------------------------------"

# === Load variables from config.env ===
Get-Content ".\config.env" -Encoding UTF8 | ForEach-Object {
    if ($_ -match "=") {
        $parts = $_ -split "=", 2
        [System.Environment]::SetEnvironmentVariable($parts[0], $parts[1])
    }
}

$VERSION        = $env:VERSION
$SERVER_PATH    = $env:SERVER_PATH
$PROJECT        = $env:PROJECT
$ICON           = $env:ICON
$MAIN_APP_FILE  = $env:MAIN_APP_FILE
$SERVER_VERSION = $env:SERVER_VERSION
$BUILD_MODE     = $env:BUILD_MODE

Write-Host "Loaded variables:"
Write-Host "VERSION: $VERSION"
Write-Host "SERVER_PATH: $SERVER_PATH"
Write-Host "PROJECT: $PROJECT"
Write-Host "ICON: $ICON"
Write-Host "MAIN_APP_FILE: $MAIN_APP_FILE"
Write-Host "SERVER_VERSION: $SERVER_VERSION"
Write-Host "------------------------------------------------"

# === Paths ===
$ASSETS_FOLDER = "assets"
# now every later use is safe
$DIST_PATH = Join-Path "$SERVER_PATH\$PROJECT\Application\versions" $VERSION
$WORK_PATH     = "build"
$APP_NAME      = "Automate $PROJECT"
$OPTIONAL_FLAGS = ""


# === Create folders if needed ===
if (!(Test-Path "$DIST_PATH")) { 
    New-Item -ItemType Directory -Path "$DIST_PATH" -Force | Out-Null 
}
if (!(Test-Path $WORK_PATH)) { New-Item -ItemType Directory -Path $WORK_PATH | Out-Null }

# === Check assets ===
if (Test-Path $ASSETS_FOLDER) {
    if ((Get-ChildItem $ASSETS_FOLDER).Count -eq 0) {
        Write-Host "Asset folder is empty. Skipping asset bundle."
    } else {
        $OPTIONAL_FLAGS += " --add-data `"$ASSETS_FOLDER;assets`""
    }
} else {
    Write-Host "Asset folder not found. Skipping asset bundle."
}

# === Check icon ===
if ($ICON -and $ICON -ne "None") {
    $OPTIONAL_FLAGS += " --icon `"$ICON`""
}

Write-Host "Detected version: $VERSION"
Write-Host "Building $APP_NAME version $VERSION..."

# === Build app ===
$cmd = "pyinstaller `"$MAIN_APP_FILE`" --name `"$APP_NAME`" $OPTIONAL_FLAGS --noconsole $BUILD_MODE --distpath `"$DIST_PATH`" --workpath `"$WORK_PATH`""
Write-Host $cmd
Invoke-Expression $cmd

Write-Host "Build completed successfully!"
Write-Host "------------------------------------------------"

# === Copy source code file into archive ===
if (Test-Path $MAIN_APP_FILE) {

    $mainFileExt  = [IO.Path]::GetExtension($MAIN_APP_FILE)
    $mainFileBase = [IO.Path]::GetFileNameWithoutExtension($MAIN_APP_FILE)

    # Build destination path relative to the PS1 script folder
    $archiveFolder = Join-Path $BaseDir "archive"

    # Create folder if it doesn't exist
    if (!(Test-Path $archiveFolder)) {
        New-Item -ItemType Directory -Path $archiveFolder | Out-Null
    }

    # Construct new filename with version suffix (your preferred order)
    $newFileName = "${mainFileBase}_${VERSION}${mainFileExt}"
    $destination = Join-Path $archiveFolder $newFileName

    # Copy the main app file into archive
    Copy-Item -Path $MAIN_APP_FILE -Destination $destination -Force

    Write-Host "Copied source file to archive as: $newFileName"

} else {
    Write-Warning "Main app file not found: $MAIN_APP_FILE"
}


# === Update version files ===
Set-Content -Path "$SERVER_PATH\$PROJECT\Application\$SERVER_VERSION" -Value $VERSION
$pastFile = ".\utils\past_versions.txt"
# Read the file as a single string
$content = Get-Content $pastFile -Raw
# If it doesn't end with a newline, add one first
if ($content -and $content[-1] -ne "`n") {
    Add-Content $pastFile ""
}
# Now safely append the version on its own line
Add-Content $pastFile $VERSION

# === Clean previous build ===
Remove-Item -Recurse -Force $WORK_PATH
Remove-Item -Force "$APP_NAME.spec" -ErrorAction SilentlyContinue

# === Final message ===
Write-Host "Version $VERSION deployed successfully to $DIST_PATH"
Pause
