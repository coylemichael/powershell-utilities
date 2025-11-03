# Teams External Access Management

Manage Teams external federation and allowed domain configurations. Implements zero-trust external collaboration by enforcing allow-list based policies.

## 📋 Scripts

### `Install-TeamsModule.ps1`
Install and validate the Microsoft Teams PowerShell module.

```powershell
.\Install-TeamsModule.ps1
```

---

### `Connect-TeamsTenant.ps1`
Establish authenticated connection to Microsoft Teams (supports MFA).

```powershell
.\Connect-TeamsTenant.ps1
```

---

### `Validate-FederationConfig.ps1`
Audit current Teams federation configuration and identify policy mode (allow-list/block-list/open/disabled).

```powershell
.\Validate-FederationConfig.ps1
```

---

### `Enable-AllowListMode.ps1`
Switch Teams federation to allow-list mode (blocks all domains until explicitly allowed).

```powershell
.\Enable-AllowListMode.ps1
```

**⚠️ WARNING**: Blocks all external federation immediately.

---

### `Manage-AllowListDomains.ps1`
Interactive menu tool for managing allowed/blocked domains. Reads from `C:\temp\domains.txt`.

```powershell
# Create domain file manually or via cli
Set-Content -Path C:\temp\domains.txt -Value "partner1.com","partner2.com"

# Run management tool
.\Manage-AllowListDomains.ps1
```

**Menu Options:**
1. Add domains from TXT file
2. Remove domains from TXT file
3. Check current allowed domains
4. Check current blocked domains
5. Check TXT file validity
6. Check connected tenant
7. Exit

---

**Author**: Michael Coyle
