@echo off
:: Batch file to run a PowerShell script as Administrator

:: Path to your PowerShell script
set "SCRIPT_PATH=%~dp0install.ps1"

:: Check if running as admin
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting administrative privileges...
    powershell -Command "Start-Process 'powershell.exe' -ArgumentList '-ExecutionPolicy Bypass -File ""%SCRIPT_PATH%""' -Verb RunAs"
    exit /b
)

:: If already running as admin, run directly
powershell -ExecutionPolicy Bypass -File "%SCRIPT_PATH%"
pause
