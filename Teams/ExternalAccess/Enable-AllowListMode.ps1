# ====================================================================================
# Script Name:   Enable-AllowListMode.ps1
# Author:        Michael Coyle, Senior Cyber Security Analyst
# Date:          30 October 2025
# Purpose:       Enable Microsoft Teams external federation and switch the tenant to
#                allow-list mode (blocking all external domains until specified).
# ====================================================================================

Write-Host "=== Enable Microsoft Teams Allow-List Mode ===" -ForegroundColor Cyan

# --- Step 1: Validate PowerShell version ---
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Error "PowerShell 5.0 or later is required. Current version: $($PSVersionTable.PSVersion)"
    exit 1
}

# --- Step 2: Check for MicrosoftTeams module ---
if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {
    Write-Error "MicrosoftTeams module not found. Please run Install-TeamsModule.ps1 first."
    exit 1
}

# --- Step 3: Import module and validate connection ---
try {
    Import-Module MicrosoftTeams -ErrorAction Stop
    Write-Host "MicrosoftTeams module imported successfully." -ForegroundColor Green
}
catch {
    Write-Error "Failed to import MicrosoftTeams module: $($_.Exception.Message)"
    exit 1
}

# Check connection (attempt if not connected)
try {
    $null = Get-CsTenant -ErrorAction Stop
    Write-Host "Connection to Teams confirmed." -ForegroundColor Green
}
catch {
    Write-Host "No active Teams session found. Attempting to connect..." -ForegroundColor Yellow
    try {
        Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
        Write-Host "Connected successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Unable to connect to Microsoft Teams: $($_.Exception.Message)"
        exit 1
    }
}

# --- Step 4: Enable federation (if disabled) and switch to allow-list mode ---
try {
    $currentConfig = Get-CsTenantFederationConfiguration -ErrorAction Stop
    if (-not $currentConfig.AllowFederatedUsers) {
        Write-Host "Enabling external federation..." -ForegroundColor Yellow
        Set-CsTenantFederationConfiguration -AllowFederatedUsers $true
    } else {
        Write-Host "Federation already enabled." -ForegroundColor Green
    }

    # Create and apply a new empty allow list
    Write-Host "Switching tenant to allow-list mode (block all until domains added)..." -ForegroundColor Yellow
    $allow = New-CsEdgeAllowList
    Set-CsTenantFederationConfiguration -AllowedDomains $allow

    Write-Host "Tenant successfully switched to allow-list mode." -ForegroundColor Green
}
catch {
    Write-Error "Failed to update federation configuration: $($_.Exception.Message)"
    exit 1
}

# --- Step 5: Validate result ---
try {
    $validation = Get-CsTenantFederationConfiguration -ErrorAction Stop |
                  Select-Object AllowFederatedUsers, AllowedDomains, AllowedDomainsAsAList, BlockedDomains

    Write-Host "`nUpdated federation configuration:" -ForegroundColor Cyan
    $validation | Format-Table -AutoSize
}
catch {
    Write-Error "Failed to retrieve updated configuration: $($_.Exception.Message)"
    exit 1
}

Write-Host "`n=== Teams allow-list mode successfully enabled ===" -ForegroundColor Cyan
Write-Host "All external domains are currently blocked until explicitly added." -ForegroundColor Yellow
Write-Host "Use Manage-AllowListDomains.ps1 to add or remove allowed domains." -ForegroundColor Gray
