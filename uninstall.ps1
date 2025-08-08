# Uninstall script for Gemini Extractor

$installPath = "C:\GeminiExtractor"

# Extensions to remove
$fileTypes = @(".rar", ".zip")

foreach ($ext in $fileTypes) {
    $baseKey = "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\$ext\shell\Gemini Extract"
    if (Test-Path $baseKey) {
        Remove-Item -Path $baseKey -Recurse -Force
        Write-Host "Removed context menu for $ext files." -ForegroundColor Green
    }
}

# Remove installed script folder
if (Test-Path $installPath) {
    Remove-Item -Path $installPath -Recurse -Force
    Write-Host "Deleted installation folder: $installPath" -ForegroundColor Green
}

Write-Host "`n=== Uninstallation completed! ===" -ForegroundColor Cyan
