# ====================================================================================
# Script Name:   Connect-TeamsTenant.ps1
# Author:        Michael Coyle, Senior Cyber Security Analyst
# Date:          30 October 2025
# Purpose:       Import the MicrosoftTeams module and establish a connection to the
#                target Microsoft Teams tenant. Includes validations for module presence,
#                import success, sign-in success, and tenant read-back.
# ====================================================================================

Write-Host "=== Connect to Microsoft Teams Tenant ===" -ForegroundColor Cyan

# --- Validation: PowerShell version (5.1+ recommended, 7.x supported) ---
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Error "PowerShell 5.0 or later is required. Current: $($PSVersionTable.PSVersion)"
    exit 1
}

# --- Validation: MicrosoftTeams module presence (install if missing) ---
$teamsModule = Get-Module -ListAvailable -Name MicrosoftTeams | Sort-Object Version -Descending | Select-Object -First 1
if (-not $teamsModule) {
    Write-Host "MicrosoftTeams module not found. Attempting install..." -ForegroundColor Yellow
    try {
        # Ensure PSGallery bits are available; install into CurrentUser scope
        if (-not (Get-Module -ListAvailable -Name PowerShellGet)) {
            Install-Module PowerShellGet -Force -AllowClobber
        }
        Install-Module MicrosoftTeams -Scope CurrentUser -Force -ErrorAction Stop
        $teamsModule = Get-Module -ListAvailable -Name MicrosoftTeams | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $teamsModule) { throw "Post-install validation failed — module still not available." }
        Write-Host "MicrosoftTeams installed. Version: $($teamsModule.Version)" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to install MicrosoftTeams module: $($_.Exception.Message)"
        Write-Host "Tip: Run Install-TeamsModule.ps1, then retry." -ForegroundColor Yellow
        exit 1
    }
} else {
    Write-Host "MicrosoftTeams found. Version: $($teamsModule.Version)" -ForegroundColor Green
}

# --- Validation: Import module ---
try {
    Import-Module MicrosoftTeams -ErrorAction Stop
    Write-Host "Module imported successfully." -ForegroundColor Green
}
catch {
    Write-Error "Failed to import MicrosoftTeams: $($_.Exception.Message)"
    exit 1
}

# --- Connect: interactive sign-in (supports MFA) ---
try {
    Write-Host "Prompting for Microsoft Teams sign-in..." -ForegroundColor Cyan
    $null = Connect-MicrosoftTeams -ErrorAction Stop
    Write-Host "Sign-in completed." -ForegroundColor Green
}
catch {
    Write-Error "Connect-MicrosoftTeams failed: $($_.Exception.Message)"
    exit 1
}

# --- Post-connection validation: read tenant context ---
try {
    $tenant = Get-CsTenant -ErrorAction Stop | Select-Object DisplayName, TenantId
    if ($null -eq $tenant -or -not $tenant.TenantId) {
        throw "Connected but unable to read tenant context (Get-CsTenant returned null)."
    }
    Write-Host "Connected to tenant:" -ForegroundColor Green
    $tenant | Format-List
}
catch {
    Write-Error "Connected, but validation failed: $($_.Exception.Message)"
    Write-Host "Check that the signed-in account has Teams Admin/Global Admin rights." -ForegroundColor Yellow
    exit 1
}

# --- Optional: quick permission probe (federation read) ---
try {
    $null = Get-CsTenantFederationConfiguration -ErrorAction Stop | Out-Null
    Write-Host "Federation configuration readable (permissions look good)." -ForegroundColor Green
}
catch {
    Write-Host "Connected, but couldn't read federation config (may be permission-scoped). Continuing..." -ForegroundColor Yellow
}

Write-Host "=== Connection established and validated ===" -ForegroundColor Cyan
# To disconnect later:  Disconnect-MicrosoftTeams
