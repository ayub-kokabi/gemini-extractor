# Ensure script is running as Administrator
if (-not ([Security.Principal.WindowsPrincipal] `
        [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Host "Please run this script as Administrator." -ForegroundColor Red
    exit 1
}

# Installation path and script name
$installPath = "C:\GeminiExtractor"
$scriptName  = "gemini-extractor.py"

# Ensure installation folder exists
if (-not (Test-Path $installPath)) {
    Write-Host "Creating installation folder at $installPath..." -ForegroundColor Cyan
    New-Item -ItemType Directory -Path $installPath | Out-Null
}

# Copy Python script to installation path
Write-Host "Copying Python script to $installPath..." -ForegroundColor Cyan
Copy-Item -Path ".\$scriptName" -Destination "$installPath\$scriptName" -Force

# Find python.exe path
$pythonPath = (Get-Command python -ErrorAction SilentlyContinue).Path
if (-not $pythonPath) {
    Write-Host "Python not found in PATH. Please install Python and try again." -ForegroundColor Red
    exit 1
}

# File extensions to register
$fileTypes = @(".rar", ".zip")

foreach ($ext in $fileTypes) {
    $baseKey = "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\$ext\shell\Extract to folder"
    $cmdKey  = "$baseKey\command"

    # Create base key if not exists
    if (-not (Test-Path $baseKey)) {
        New-Item -Path $baseKey -Force | Out-Null
    }
    Set-ItemProperty -Path $baseKey -Name "(default)" -Value "Extract to folder"

    # Create command key
    if (-not (Test-Path $cmdKey)) {
        New-Item -Path $cmdKey -Force | Out-Null
    }
    Set-ItemProperty -Path $cmdKey -Name "(default)" `
        -Value "`"$pythonPath`" `"$installPath\$scriptName`" `"%1`""

    Write-Host "Context menu option for $ext added successfully." -ForegroundColor Green
}

Write-Host "Installation completed!" -ForegroundColor Green
