# ====================================================================================
# Script Name:   Install-TeamsModule.ps1
# Author:        Michael Coyle, Senior Cyber Security Analyst
# Date:          30 October 2025
# Purpose:       Install the Microsoft Teams PowerShell module from PSGallery.
#                Checks for existing installations before attempting new installs.
# ====================================================================================

Write-Host "=== Microsoft Teams PowerShell Module Setup ===" -ForegroundColor Cyan

# --- Step 1: Ensure script execution policy allows module installation ---
$currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
if ($currentPolicy -ne "RemoteSigned") {
    Write-Host "Setting execution policy to RemoteSigned for CurrentUser scope..." -ForegroundColor Yellow
    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
} else {
    Write-Host "Execution policy already set to RemoteSigned." -ForegroundColor Green
}

# --- Step 2: Check for PowerShellGet ---
if (-not (Get-Module -ListAvailable -Name PowerShellGet)) {
    Write-Host "PowerShellGet not found. Installing..." -ForegroundColor Yellow
    Install-Module PowerShellGet -Force -AllowClobber
} else {
    Write-Host "PowerShellGet is already installed." -ForegroundColor Green
}

# --- Step 3: Check for MicrosoftTeams module ---
if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {
    Write-Host "MicrosoftTeams module not found. Installing..." -ForegroundColor Yellow
    Install-Module MicrosoftTeams -Scope CurrentUser -Force
} else {
    Write-Host "MicrosoftTeams module is already installed." -ForegroundColor Green
}

# --- Step 4: Validate installation ---
if (Get-Module -ListAvailable -Name MicrosoftTeams) {
    Write-Host "MicrosoftTeams module successfully validated." -ForegroundColor Green
    $teamsModule = Get-Module -ListAvailable -Name MicrosoftTeams | Sort-Object Version -Descending | Select-Object -First 1
    Write-Host ("Installed version: " + $teamsModule.Version)
} else {
    Write-Error "MicrosoftTeams module installation failed. Please check your network or PSGallery access."
}

Write-Host "=== Module setup completed ===" -ForegroundColor Cyan
