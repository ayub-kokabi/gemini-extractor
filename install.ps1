# Ensure script is running as Administrator
if (-not ([Security.Principal.WindowsPrincipal] `
        [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Host "❌ Please run this script as Administrator." -ForegroundColor Red
    exit 1
}

function Write-Section {
    param([string]$message)
    Write-Host "`n======================================" -ForegroundColor DarkGray
    Write-Host $message -ForegroundColor Cyan
    Write-Host "======================================" -ForegroundColor DarkGray
}

Write-Host "`n=== Gemini Extractor Installer ===`n" -ForegroundColor Magenta

# Paths
$installPath = "C:\GeminiExtractor"
$scriptName  = "gemini-extractor.py"
$iconPath    = "$installPath\icon.ico"

# 1. Check Python installation
Write-Section "[1/6] Checking Python installation..."
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "❌ Python is not installed or not in PATH. Please install Python first." -ForegroundColor Red
    exit 1
}
Write-Host "✅ Python found." -ForegroundColor Green
$pythonPath = (Get-Command python).Source

# 2. Install required Python packages
Write-Section "[2/6] Installing required Python packages..."
$packages = @("pyzipper", "rarfile", "google-genai", "rich")
foreach ($pkg in $packages) {
    Write-Host "📦 Installing $pkg..." -ForegroundColor Yellow
    python -m pip install --upgrade $pkg --quiet --disable-pip-version-check *> $null
}
Write-Host "✅ All packages installed." -ForegroundColor Green

# 3. Check WinRAR installation
Write-Section "[3/6] Checking WinRAR installation..."
try {
    $winrarPath = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WinRAR.exe")."(default)"
} catch {
    $winrarPath = $null
}
if (-not $winrarPath -or -not (Test-Path $winrarPath)) {
    Write-Host "❌ WinRAR is not installed or not found in registry. Please install WinRAR first." -ForegroundColor Red
    exit 1
}
Write-Host "✅ WinRAR found at: $winrarPath" -ForegroundColor Green

# 4. Check GEMINI_API_KEY
Write-Section "[4/6] Checking GEMINI_API_KEY..."
$apiKey = [System.Environment]::GetEnvironmentVariable("GEMINI_API_KEY", "User")
if ($apiKey) {
    Write-Host "✅ GEMINI_API_KEY already set." -ForegroundColor Green
} else {
    Write-Host "⚠️ No GEMINI_API_KEY found." -ForegroundColor Red
    $openKeyPage = Read-Host "Get a free API Key from Google AI Studio? (y/n)"
    if ($openKeyPage -match "^[Yy]") {
        Start-Process "https://aistudio.google.com/apikey"
    }
    $newKey = Read-Host "Enter your GEMINI_API_KEY"
    if (-not $newKey) {
        Write-Host "❌ No API key entered. Exiting..." -ForegroundColor Red
        exit 1
    }
    [System.Environment]::SetEnvironmentVariable("GEMINI_API_KEY", $newKey, "User")
    $env:GEMINI_API_KEY = $newKey
    Write-Host "✅ GEMINI_API_KEY saved successfully." -ForegroundColor Green
}

# 5. Copy files to installation folder
Write-Section "[5/6] Preparing installation folder..."
if (-not (Test-Path $installPath)) {
    Write-Host "📂 Creating installation folder at $installPath..." -ForegroundColor Cyan
    New-Item -ItemType Directory -Path $installPath | Out-Null
}
Write-Host "📄 Copying Python script..." -ForegroundColor Cyan
Copy-Item -Path ".\$scriptName" -Destination "$installPath\$scriptName" -Force

if (Test-Path ".\icon.ico") {
    Copy-Item -Path ".\icon.ico" -Destination $iconPath -Force
}

Write-Host "✅ Files copied successfully." -ForegroundColor Green

# 6. Add context menu
Write-Section "[6/6] Adding context menu options..."
$fileTypes = @(".rar", ".zip")

foreach ($ext in $fileTypes) {
    $baseKey = "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\$ext\shell\Gemini Extract"
    $cmdKey  = "$baseKey\command"

    if (-not (Test-Path $baseKey)) { New-Item -Path $baseKey -Force | Out-Null }
    Set-ItemProperty -Path $baseKey -Name "(default)" -Value "Gemini Extract"

    if (Test-Path $iconPath) {
        Set-ItemProperty -Path $baseKey -Name "Icon" -Value "`"$iconPath`""
    }

    if (-not (Test-Path $cmdKey)) { New-Item -Path $cmdKey -Force | Out-Null }
    Set-ItemProperty -Path $cmdKey -Name "(default)" `
        -Value "`"$pythonPath`" `"$installPath\$scriptName`" `"%1`""

    Write-Host "✅ Context menu option for $ext added." -ForegroundColor Green
}

Write-Host "`n🎉 Installation complete! You can now right-click a .rar or .zip file to extract with Gemini Extractor." -ForegroundColor Magenta
