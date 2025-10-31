# ====================================================================================
# Script Name:   Validate-FederationConfig.ps1
# Author:        Michael Coyle, Senior Cyber Security Analyst
# Date:          30 October 2025
# Purpose:       Validate the current Microsoft Teams federation configuration.
#                Detects active mode (allow-list, block-list, open, disabled, or empty)
#                and outputs a concise, color-coded summary.
# ====================================================================================

Write-Host "=== Validating Microsoft Teams Federation Configuration ===" -ForegroundColor Cyan

# --- Step 1: PowerShell version check ---
if ($PSVersionTable.PSVersion.Major -lt 5) { Write-Error "PowerShell 5.0 or later is required."; exit 1 }

# --- Step 2: Ensure module exists ---
if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {
    Write-Error "MicrosoftTeams module not found. Please run Install-TeamsModule.ps1 first."
    exit 1
}

# --- Step 3: Import module ---
try {
    Import-Module MicrosoftTeams -ErrorAction Stop
    Write-Host "MicrosoftTeams module imported successfully." -ForegroundColor Green
} catch {
    Write-Error "Import failed: $($_.Exception.Message)"
    exit 1
}

# --- Step 4: Validate Teams connection ---
try {
    $null = Get-CsTenant -ErrorAction Stop
    Write-Host "Connection to Teams tenant confirmed." -ForegroundColor Green
} catch {
    Write-Host "No active Teams session found. Attempting to connect..." -ForegroundColor Yellow
    try {
        Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
        Write-Host "Connected successfully." -ForegroundColor Green
    } catch {
        Write-Error "Failed to connect to Microsoft Teams: $($_.Exception.Message)"
        exit 1
    }
}

# --- Step 5: Retrieve federation configuration ---
try { 
    $config = Get-CsTenantFederationConfiguration -ErrorAction Stop 
} catch { 
    Write-Error "Failed to retrieve configuration: $($_.Exception.Message)"
    exit 1 
}

# --- Display configuration ---
Write-Host "`n--- Federation Configuration ---" -ForegroundColor Cyan
$config | Select-Object AllowFederatedUsers, AllowedDomains, AllowedDomainsAsAList, BlockedDomains | Format-List

# --- Step 6: Determine policy mode ---
Write-Host "--- Federation Summary ---" -ForegroundColor Cyan

if (-not $config.AllowFederatedUsers) {
    Write-Host "Federation: DISABLED — no external access permitted." -ForegroundColor Red
}
elseif ($config.AllowedDomainsAsAList -and $config.AllowedDomainsAsAList.Count -gt 0) {
    Write-Host ("Federation: ALLOW-LIST MODE — " + $config.AllowedDomainsAsAList.Count + " domain(s) allowed.") -ForegroundColor Green
}
elseif ($config.BlockedDomains -and $config.BlockedDomains.Count -gt 0) {
    Write-Host ("Federation: BLOCK-LIST MODE — " + $config.BlockedDomains.Count + " domain(s) blocked.") -ForegroundColor Yellow
}
elseif ($config.AllowedDomains -eq "AllowAllKnownDomains") {
    Write-Host "Federation: OPEN — all known domains can federate." -ForegroundColor Yellow
}
else {
    Write-Host "Federation: ENABLED but NO domains configured — all external domains currently BLOCKED." -ForegroundColor Red
}

# --- Step 7: Check Teams consumer access ---
try {
    if ($config.AllowTeamsConsumer -or $config.AllowTeamsConsumerInbound) {
        Write-Host "Consumer Access: ENABLED (personal Microsoft accounts can federate)." -ForegroundColor Yellow
    } else {
        Write-Host "Consumer Access: DISABLED (organization-only federation)." -ForegroundColor Green
    }
} catch {
    Write-Host "Consumer Access: Unknown — unable to determine setting." -ForegroundColor Yellow
}

Write-Host "Validation complete." -ForegroundColor Cyan
