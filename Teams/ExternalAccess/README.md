# Teams External Access Management

Manage Teams external federation and allowed domain configurations. Implements zero-trust external collaboration by enforcing allow-list based policies.

## Scripts

### Install-TeamsModule.ps1
Install and validate the Microsoft Teams PowerShell module.

### Connect-TeamsTenant.ps1
Establish authenticated connection to Microsoft Teams (supports MFA).

### Validate-FederationConfig.ps1
Audit current Teams federation configuration and identify policy mode (allow-list/block-list/open/disabled).

### Enable-AllowListMode.ps1
Switch Teams federation to allow-list mode (blocks all domains until explicitly allowed).  
**⚠️ WARNING**: Blocks all external federation immediately.

### Manage-AllowListDomains.ps1
Interactive menu tool for managing allowed/blocked domains. Reads from `C:\temp\domains.txt`.

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
