# AppReport

Read-only inventory of Entra ID app registrations and enterprise apps with access to Microsoft 365 workloads, with focus on mail.

| Area | Content |
|---|---|
| Permissions | Configured (manifest) and granted (admin consent / app role assignment / user consent) permissions on Microsoft Graph, Exchange Online, SharePoint Online, Office 365 Management APIs, Teams Tenant Admin API |
| Scope | Graph app permissions correlated with **Application Access Policies** and **Exchange RBAC for Applications** (management scope / administrative unit), grouped by target user / group / scope |
| Credentials | Secrets and certificates of app registration and service principal, `Expired` / `Expiring` / `Valid`, separate tabs |
| Categories | One tab per app type: Exchange, SharePoint, Teams, Reporting, MailStore, Backup, Salesforce, OutlookOffice, Other |
| EWS | All `Ews*` parameters of `Get-OrganizationConfig` plus mailbox-level exceptions (`EwsEnabled`, `EwsApplicationAccessPolicy`, `EwsAllowList`, `EwsBlockList`, …) |
| Export | CSV per sheet (`;`-delimited, UTF-8), optional XLSX workbook via ImportExcel |

## Requirements

| Module | Purpose |
|---|---|
| `Microsoft.Graph.Authentication` | Graph connection and requests (mandatory) |
| `ExchangeOnlineManagement` 3.1+ | Access policies, RBAC, EWS config (optional) |
| `ImportExcel` | XLSX output (optional) |

```powershell
Install-Module Microsoft.Graph.Authentication, ExchangeOnlineManagement, ImportExcel -Scope CurrentUser
Import-Module .\AppReport\AppReport.psd1
```

### Permissions

| Mode | Graph | Exchange Online |
|---|---|---|
| Interactive (delegated) | `Application.Read.All`, `Directory.Read.All`, `DelegatedPermissionGrant.Read.All`, role Global Reader | Global Reader or View-Only Organization Management |
| App-only (secret / certificate) | Application permissions `Application.Read.All`, `Directory.Read.All` | `Exchange.ManageAsApp` + Entra role Global Reader (or EXO role group View-Only Organization Management) |

App-only Exchange with **client secret** uses `Connect-ExchangeOnline -AccessToken` (client-credentials token for `https://outlook.office365.com/.default`). Certificate auth is the Microsoft-documented path for unattended EXO.

## Functions

| Function | Purpose |
|---|---|
| `Set-AppReportConfig` | Saves TenantId / ClientId / Organization and either a DPAPI-protected secret or a certificate thumbprint as JSON. Default `C:\temp\config\AppReport.config.json`; folder is created when missing |
| `Get-AppReportConfig` | Reads the JSON and returns a splat hashtable (secret as SecureString) |
| `Connect-AppReport` | Graph (+ `-Exchange`) connection: interactive, client secret, certificate, or `-ConfigPath` |
| `Get-AppReportInventory` | Graph inventory: `Permissions`, `Credentials`, `Apps` |
| `Get-AppReportAccessPolicy` | Application Access Policies, scope group resolved, `-ExpandMembers` for nested membership |
| `Get-AppReportExchangeRbac` | RBAC for Applications role assignments incl. management scope filter / AU |
| `Get-AppReportEwsConfig` | EWS organization parameters and mailbox exceptions |
| `Get-AppReportScope` | Correlation of all sources, one row per target × app |
| `Invoke-AppReport` | Runs everything, writes CSV / XLSX, returns result object |

## Workbook tabs

| Tab | Content |
|---|---|
| `Summary` | One row per app: category, permission categories, application/delegated permissions, UsesEWS, ScopeType, NextExpiry, CredentialState, AccessPolicy, ExchangeRbacScope, EffectiveMailScope |
| `Permissions` | One row per app × resource × permission × evidence |
| `Credentials` | All secrets / certificates with EndDate, DaysLeft, State |
| `Expired` / `Expiring` | Credentials with State `Expired` / ending within `-ExpiringDays` (default 30) |
| `Exchange`, `SharePoint`, `Teams`, `Reporting`, `MailStore`, `Backup`, `Salesforce`, `OutlookOffice`, `Other` | Summary rows per app category |
| `ScopeByTarget` | Target (mailbox, group member, group, management scope, `ALL MAILBOXES`, consenting user) → app, source, restriction, permissions |
| `AccessPolicies` | Raw Application Access Policies (members expanded with `-ExpandAccessPolicyMembers`) |
| `ExchangeRBAC` | RBAC for Applications assignments with scope |
| `EWSConfig` | Org `Ews*` settings and mailbox exceptions |

Empty sheets are not written.

### Field values

| Field | Values |
|---|---|
| `Evidence` | `Configured` = requested in manifest · `Granted` = consent / assignment exists |
| `PermissionCategory` | EWS, ExchangeAdmin, Mail, Calendar, Contacts, SharePoint, Teams, Reporting, Directory, Other |
| `ScopeType` | Tenant-wide (all mailboxes) · Selected resources · Exchange admin (role-based) · Single user (delegated) · All users (delegated, signed-in context) |
| `CredentialState` | Valid · Expiring · PartlyExpired · AllExpired · NoCredential (local app without credential) · External/Managed (multi-tenant or managed identity, credentials live in the publisher tenant) |
| `EffectiveMailScope` | Tenant-wide (unrestricted) · Restricted (ApplicationAccessPolicy) · Tenant-wide minus denied |

### App category

Categories are assigned by display name / publisher (regex map), then by dominant permission workload. Extend or override with `-AppCategoryMap` (entries are evaluated first):

```powershell
$map = [ordered]@{ Backup = 'Vembu|Hycu'; Reporting = 'Contoso Reporting' }
Invoke-AppReport -AppCategoryMap $map -Xlsx
```

## Examples

### 1. Interactive, full report with XLSX

```powershell
Import-Module .\AppReport\AppReport.psd1
Invoke-AppReport -TenantId contoso.onmicrosoft.com -Exchange -Xlsx -OutputPath C:\temp\AppReport
```

### 2. Store client secret once, run unattended

```powershell
Set-AppReportConfig -TenantId contoso.onmicrosoft.com `
                    -ClientId 11111111-2222-3333-4444-555555555555 `
                    -ClientSecret (Read-Host -AsSecureString 'Client secret')
# -> C:\temp\config\AppReport.config.json (folder created if missing)

Invoke-AppReport -ConfigPath C:\temp\config\AppReport.config.json -Exchange -Xlsx
```

The JSON is a splat; `ClientSecret` is DPAPI-encrypted and can only be decrypted by the same user on the same machine:

```json
{
  "TenantId": "contoso.onmicrosoft.com",
  "ClientId": "11111111-2222-3333-4444-555555555555",
  "Organization": "contoso.onmicrosoft.com",
  "AuthType": "Secret",
  "ClientSecret": "01000000d08c9ddf0115d1118c7a00c04fc297eb..."
}
```

```powershell
$cfg = Get-AppReportConfig -Path C:\temp\config\AppReport.config.json
Connect-AppReport @cfg -Exchange
```

### 3. Certificate

```powershell
Set-AppReportConfig -TenantId contoso.onmicrosoft.com -ClientId $appId `
                    -CertificateThumbprint 'A1B2C3D4E5F6...' -Organization contoso.onmicrosoft.com `
                    -Path C:\temp\config\contoso.json

Invoke-AppReport -ConfigPath C:\temp\config\contoso.json -Exchange -Xlsx
# or directly
Invoke-AppReport -TenantId $tid -ClientId $appId -CertificateThumbprint $thumb -Organization contoso.onmicrosoft.com -Exchange -Xlsx
```

### 4. Existing sessions

```powershell
Connect-MgGraph -Scopes Application.Read.All,Directory.Read.All -NoWelcome
Connect-ExchangeOnline -ShowBanner:$false
$r = Invoke-AppReport -Xlsx
```

### 5. Unrestricted mail apps (no Application Access Policy, no RBAC scope)

```powershell
$r.Apps | Where-Object EffectiveMailScope -eq 'Tenant-wide (unrestricted)' |
    Format-Table AppName, AppCategory, ApplicationPerms, CredentialState
```

### 6. Who can access a specific mailbox / group

```powershell
$r.ScopeByTarget | Group-Object Target | Sort-Object Count -Descending | Select-Object Count, Name
$r.ScopeByTarget | Where-Object Target -eq 'scanner@contoso.com' | Format-Table AppName, Source, Restriction, Permissions
```

Group scopes are expanded to members with `-ExpandAccessPolicyMembers` (rows `TargetType = User (via group)`, `ViaGroup = <group>`).

### 7. Expired and expiring credentials

```powershell
Invoke-AppReport -ExpiringDays 60 -NoCsv | Select-Object -ExpandProperty Credentials |
    Where-Object State -in 'Expired','Expiring' | Sort-Object EndDate |
    Format-Table AppName, OwnerType, CredentialType, DisplayName, EndDate, DaysLeft, State
```

### 8. EWS: apps using EWS vs. org / mailbox exceptions

```powershell
$r.Apps      | Where-Object UsesEWS | Format-Table AppName, ApplicationPerms, DelegatedPerms, ExchangeRbacScope
$r.EwsConfig | Where-Object Level -eq 'Organization' | Format-Table Setting, Value
$r.EwsConfig | Where-Object Exception | Format-Table PrimarySmtpAddress, Setting, Value, OrgValue
```

### 9. Single building blocks

```powershell
Connect-AppReport -TenantId contoso.onmicrosoft.com -Exchange
$inv  = Get-AppReportInventory -ExcludeMicrosoftFirstParty
$aap  = Get-AppReportAccessPolicy -ExpandMembers
$rbac = Get-AppReportExchangeRbac
$ews  = Get-AppReportEwsConfig -SkipMailboxes
$scope = Get-AppReportScope -Inventory $inv -AccessPolicy $aap -ExchangeRbac $rbac
$scope | Export-Csv .\ScopeByTarget.csv -NoTypeInformation -Delimiter ';' -Encoding UTF8
```

### 10. Graph only, all permissions (incl. Directory / User.Read)

```powershell
Invoke-AppReport -TenantId contoso.onmicrosoft.com -AllPermissions -Xlsx
```

## Parameters of Invoke-AppReport

| Parameter | Default | Description |
|---|---|---|
| `-TenantId`, `-ClientId`, `-ClientSecret`, `-CertificateThumbprint`, `-Organization` | – | Connection (see Connect-AppReport) |
| `-ConfigPath` | – | JSON from Set-AppReportConfig |
| `-Exchange` | off | Connect Exchange Online as well |
| `-OutputPath` | `C:\temp\AppReport` | Created when missing |
| `-Xlsx` | off | Workbook via ImportExcel |
| `-NoCsv` | off | Skip CSV files |
| `-AllPermissions` | off | Include Directory/User/Other permissions |
| `-ExcludeMicrosoftFirstParty` | off | Skip Microsoft-owned service principals |
| `-ExpiringDays` | 30 | Threshold for `Expiring` |
| `-ExpandAccessPolicyMembers` | off | One row per member of the scope group |
| `-SkipExchangeRbac`, `-SkipEws`, `-SkipEwsMailboxes` | off | Skip parts (mailbox scan is the slow part in large tenants) |
| `-AppCategoryMap` | – | Additional / overriding name→category regex map |

## Interpretation notes

* `Configured` permissions are requests in the manifest, not consent. `Granted` permissions exist, but do not prove use (check sign-in logs).
* Graph application permissions for Mail/Calendar/Contacts/EWS without an Application Access Policy or RBAC scope apply to **all mailboxes**.
* Delegated permissions are limited by the signed-in user's own access.
* Multi-tenant apps (`External/Managed`) have their credentials in the publisher tenant; the expiry columns stay empty.
* SharePoint `Sites.Selected` per-site grants are not enumerated (scope shown as `SELECTED RESOURCES`).
* Not covered: SMTP relay connectors, mailbox folder / FullAccess ACLs, sign-in activity.
