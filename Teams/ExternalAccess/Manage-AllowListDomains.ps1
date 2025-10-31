# ====================================================================================
# Script Name:   Manage-AllowListDomains.ps1
# Author:        Michael Coyle, Senior Cyber Security Analyst
# Date:          31 October 2025
# Purpose:       Menu-driven management of Teams federation allow/blocked domains.
#                1) Add  2) Remove  3) Check allowed  4) Check blocked
#                5) Check TXT file  6) Check tenant  7) Exit
# ====================================================================================

$DomainFile = "C:\temp\domains.txt"   # one domain per line

# ---------- Basic checks ----------
if ($PSVersionTable.PSVersion.Major -lt 5) { Write-Error "PowerShell 5.0+ required."; exit 1 }
if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) { Write-Error "MicrosoftTeams module not found. Run Install-TeamsModule.ps1."; exit 1 }

try { Import-Module MicrosoftTeams -ErrorAction Stop; Write-Host "MicrosoftTeams module imported successfully." -ForegroundColor Green }
catch { Write-Error "Import failed: $($_.Exception.Message)"; exit 1 }

try { $null = Get-CsTenant -ErrorAction Stop; Write-Host "Connection to Teams tenant confirmed." -ForegroundColor Green }
catch {
  Write-Host "No active Teams session. Connecting..." -ForegroundColor Yellow
  try { Connect-MicrosoftTeams -ErrorAction Stop | Out-Null; Write-Host "Connected successfully." -ForegroundColor Green }
  catch { Write-Error "Failed to connect: $($_.Exception.Message)"; exit 1 }
}

# ---------- Helpers ----------
function Get-DomainsFromTxt {
  param([string]$Path)
  if (-not (Test-Path $Path)) { throw "File not found: $Path" }
  $raw = Get-Content $Path | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
  $pattern = '^(?!-)(?:[a-zA-Z0-9-]{1,63}(?<!-)\.)+[A-Za-z]{2,63}$'
  @{
    Valid   = $raw | Where-Object { $_ -match $pattern } | Select-Object -Unique
    Invalid = $raw | Where-Object { $_ -notmatch $pattern }
  }
}

function Confirm-Action($Prompt){ do { $c=(Read-Host "$Prompt [y/n]").ToLower().Trim() } until($c -in "y","n"); $c -eq "y" }

# Parse allowed domains from config → returns clean list like "accelins.com"
function Get-AllowedDomainsFromConfig {
  try { $cfg = Get-CsTenantFederationConfiguration -ErrorAction Stop } catch { return @() }
  if ($cfg.AllowedDomainsAsAList) { return @($cfg.AllowedDomainsAsAList) }
  if ($cfg.AllowedDomains -and $cfg.AllowedDomains -ne "AllowAllKnownDomains") {
    if ($cfg.AllowedDomains -is [string]) {
      $m = [regex]::Matches($cfg.AllowedDomains,'Domain=([^,]+)')
      if ($m.Count -gt 0) { return ($m | ForEach-Object { $_.Groups[1].Value }) }
    }
    if ($cfg.AllowedDomains.PSObject.Properties.Name -contains 'AllowedDomain') {
      return (@($cfg.AllowedDomains.AllowedDomain) -as [string[]]) -replace '^Domain='
    }
    if ($cfg.AllowedDomains.PSObject.Properties.Name -contains 'Domains') {
      return (@($cfg.AllowedDomains.Domains) -as [string[]]) -replace '^Domain='
    }
  }
  @()
}

# Clean poll with spinner (10s) — timeout message always on a NEW line
function Wait-ForAllowedDomainsUpdate {
  param([int]$MaxSeconds = 10, [int]$IntervalSeconds = 2)
  $deadline = (Get-Date).AddSeconds($MaxSeconds)
  $last = (Get-AllowedDomainsFromConfig | Sort-Object) -join ','
  $spin = @('|','/','-','\'); $i = 0
  Write-Host ("Polling for configuration updates (max {0}s)..." -f $MaxSeconds) -ForegroundColor Yellow
  Write-Host -NoNewline "Waiting for Teams to propagate "
  do {
    Start-Sleep -Seconds $IntervalSeconds
    $current = (Get-AllowedDomainsFromConfig | Sort-Object) -join ','
    if ($current -ne $last) { Write-Host "`rConfiguration updated successfully!                    " -ForegroundColor Green; return }
    Write-Host -NoNewline ("`b{0}" -f $spin[$i % $spin.Length]); $i++
  } while ((Get-Date) -lt $deadline)
  Write-Host ""
  Write-Host ("Polling finished after {0} seconds — results may still be propagating." -f $MaxSeconds) -ForegroundColor Yellow
}

# ---------- Menu ----------
Write-Host "`n=== Manage Teams External Access ===" -ForegroundColor Cyan
Write-Host "`n1) Add domains from TXT file"
Write-Host "2) Remove domains from TXT file"
Write-Host "3) Check current allowed domains"
Write-Host "4) Check current blocked domains"
Write-Host "5) Check TXT file for valid/invalid domains"
Write-Host "6) Check connected tenant (DisplayName & TenantId)"
Write-Host "7) Exit"

# ---------- Loop ----------
do {
  $choice = (Read-Host "`nEnter choice (1-7)").Trim()

  switch ($choice) {

    "1" { # Add
      try { $loaded = Get-DomainsFromTxt -Path $DomainFile } catch { Write-Error $_; continue }
      $desired = $loaded.Valid; if ($desired.Count -eq 0) { Write-Host "No valid domains found to add."; continue }
      $existing = Get-AllowedDomainsFromConfig
      $toAdd = $desired | Where-Object { $_ -notin $existing }
      $skipped = $desired | Where-Object { $_ -in $existing }

      Write-Host ""
      Write-Host ("Ready to ADD {0} domain(s)." -f $toAdd.Count) -ForegroundColor Cyan
      if ($skipped.Count -gt 0) { Write-Host ("(Skipping {0} already present: {1})" -f $skipped.Count, ($skipped -join ", ")) -ForegroundColor DarkGray }
      if ($toAdd.Count -eq 0) { Write-Host "Nothing to add." -ForegroundColor Yellow; continue }

      Write-Host ""
      if (Confirm-Action "Confirm add") {
        Write-Host ""
        try {
          Set-CsTenantFederationConfiguration -AllowedDomainsAsAList @{ Add = $toAdd } -ErrorAction Stop
          Write-Host "Domains added successfully." -ForegroundColor Green
          Wait-ForAllowedDomainsUpdate -MaxSeconds 10
          Write-Host ""
          $post = Get-AllowedDomainsFromConfig
          if ($post.Count -gt 0) {
            $joined = ($post | Sort-Object) -join ", "
            Write-Host ("Confirmed ({0}) total AllowedDomains: {1}" -f $post.Count, $joined) -ForegroundColor Cyan
          } else {
            Write-Host "Confirmed (0) total AllowedDomains: (none)" -ForegroundColor Yellow
          }
        } catch { Write-Error "Failed to add domains: $($_.Exception.Message)" }
      } else { Write-Host "Add cancelled." -ForegroundColor Yellow }
    }

    "2" { # Remove
      try { $loaded = Get-DomainsFromTxt -Path $DomainFile } catch { Write-Error $_; continue }
      $desired = $loaded.Valid; if ($desired.Count -eq 0) { Write-Host "No valid domains found to remove."; continue }
      $existing = Get-AllowedDomainsFromConfig
      $toRemove = $desired | Where-Object { $_ -in $existing }
      $skipped  = $desired | Where-Object { $_ -notin $existing }

      Write-Host ""
      Write-Host ("Ready to REMOVE {0} domain(s)." -f $toRemove.Count) -ForegroundColor Cyan
      if ($skipped.Count -gt 0) { Write-Host ("(Skipping {0} not found: {1})" -f $skipped.Count, ($skipped -join ", ")) -ForegroundColor DarkGray }
      if ($toRemove.Count -eq 0) { Write-Host "Nothing to remove." -ForegroundColor Yellow; continue }

      Write-Host ""
      if (Confirm-Action "Confirm remove") {
        Write-Host ""
        try {
          Set-CsTenantFederationConfiguration -AllowedDomainsAsAList @{ Remove = $toRemove } -ErrorAction Stop
          Write-Host "Domains removed successfully." -ForegroundColor Red
          Wait-ForAllowedDomainsUpdate -MaxSeconds 10
          Write-Host ""
          $post = Get-AllowedDomainsFromConfig
          if ($post.Count -gt 0) {
            $joined = ($post | Sort-Object) -join ", "
            Write-Host ("Confirmed ({0}) total AllowedDomains: {1}" -f $post.Count, $joined) -ForegroundColor Cyan
          } else {
            Write-Host "Confirmed (0) total AllowedDomains: (none)" -ForegroundColor Yellow
          }
        } catch { Write-Error "Failed to remove domains: $($_.Exception.Message)" }
      } else { Write-Host "Remove cancelled." -ForegroundColor Yellow }
    }

    "3" {
      Write-Host ""
      $list = Get-AllowedDomainsFromConfig
      if ($list.Count -gt 0) {
        Write-Host ("AllowedDomains: {0}" -f (($list | Sort-Object) -join ", ")) -ForegroundColor Cyan
      } else {
        Write-Host "AllowedDomains: (none)" -ForegroundColor Yellow
      }
    }

    "4" {
      Write-Host ""
      try {
        $cfg = Get-CsTenantFederationConfiguration -ErrorAction Stop
        $blocked = @($cfg.BlockedDomains)
        if ($blocked -and $blocked.Count -gt 0) {
          Write-Host ("BlockedDomains: {0}" -f (($blocked | Sort-Object) -join ", ")) -ForegroundColor Cyan
        } else {
          Write-Host "BlockedDomains: (none)" -ForegroundColor Yellow
        }
      } catch { Write-Error "Failed to read blocked domains: $($_.Exception.Message)" }
    }

    "5" {
      try { $loaded = Get-DomainsFromTxt -Path $DomainFile } catch { Write-Error $_; continue }
      if ($loaded.Valid.Count -gt 0) {
        Write-Host ""
        $joined = ($loaded.Valid | Sort-Object) -join ", "
        Write-Host ("Valid TXT domains ({0}): {1}" -f $loaded.Valid.Count, $joined) -ForegroundColor Cyan
      } else {
        Write-Host "`nNo valid domains found." -ForegroundColor Yellow
      }
    }

    "6" {
      Write-Host "`n------------------------------------------------------------" -ForegroundColor Cyan
      Write-Host " Current Connected Tenant" -ForegroundColor Cyan
      Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
      $t = Get-CsTenant -ErrorAction Stop | Select-Object DisplayName, TenantId
      Write-Host (" Tenant: {0}" -f $t.DisplayName) -ForegroundColor Green
      Write-Host (" Tenant ID: {0}" -f $t.TenantId) -ForegroundColor Yellow
      Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
    }

    "7" { Write-Host "`nExiting..." -ForegroundColor Cyan }

    default { if ($choice -notin "1","2","3","4","5","6","7") { Write-Host "Invalid selection. Choose 1–7." -ForegroundColor Yellow } }
  }

} until ($choice -eq "7")
