# Uninstall script for Gemini Extractor

# Installation path
$installPath = "C:\GeminiExtractor"
$scriptName = "gemini-extractor.py"

Write-Host "Starting uninstallation..." -ForegroundColor Yellow

# Remove context menu for RAR
$rarKey = "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.rar\shell\Extract to folder"
if (Test-Path $rarKey) {
    Remove-Item -Path $rarKey -Recurse -Force
    Write-Host "Removed context menu for .rar files." -ForegroundColor Green
}

# Remove context menu for ZIP
$zipKey = "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.zip\shell\Extract to folder"
if (Test-Path $zipKey) {
    Remove-Item -Path $zipKey -Recurse -Force
    Write-Host "Removed context menu for .zip files." -ForegroundColor Green
}

# Remove installed script folder
if (Test-Path $installPath) {
    Remove-Item -Path $installPath -Recurse -Force
    Write-Host "Deleted installation folder: $installPath" -ForegroundColor Green
}

Write-Host "Uninstallation completed!" -ForegroundColor Cyan
Pause
