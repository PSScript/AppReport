#Requires -Version 5.1
<#
    AppReport.psm1
    Inventory of mail- and M365-workload-related Entra app registrations / enterprise apps,
    their permissions, credential expiry, Exchange Application Access Policies,
    Exchange RBAC for Applications and the EWS organization / mailbox exceptions.

    Read-only. Commercial cloud (Global) only.
    Repository : PSScript/AppReport
#>

Set-StrictMode -Version 1.0

#region ---- constants ---------------------------------------------------------

$script:GraphBase       = 'https://graph.microsoft.com/v1.0'
$script:DefaultConfig   = 'C:\temp\config\AppReport.config.json'
$script:GraphScopes     = @('Application.Read.All', 'Directory.Read.All', 'DelegatedPermissionGrant.Read.All')

# Resource APIs (first-party) whose permissions are evaluated
$script:ResourceApis = [ordered]@{
    '00000003-0000-0000-c000-000000000000' = 'Microsoft Graph'
    '00000002-0000-0ff1-ce00-000000000000' = 'Office 365 Exchange Online'
    '00000003-0000-0ff1-ce00-000000000000' = 'Office 365 SharePoint Online'
    'c5393580-f805-4401-95e8-94b7a6ef2fc2' = 'Office 365 Management APIs'
    '48ac35b8-9aa8-4d74-927d-1f4a14a0b239' = 'Skype and Teams Tenant Admin API'
}

# Microsoft-owned tenants (appOwnerOrganizationId) => first-party app
$script:MicrosoftOwnerTenants = @('f8cdef31-a31e-4b4a-93e4-5f571e91255a', '72f988bf-86f1-41af-91ab-2d7cd011db47')

# Permission value => workload category. First match wins.
$script:PermissionCategories = [ordered]@{
    'EWS'           = '^(full_access_as_app|EWS\.AccessAsUser\.All|EWS\.)$|^EWS\.'
    'ExchangeAdmin' = '^(Exchange\.Manage|Exchange\.ManageAsApp|Exchange\.ManageV2|Exchange\.ManageAsAppV2)$'
    'Mail'          = '^(Mail\.|Mailbox|MailboxSettings\.|MailboxFolder\.|MailboxItem\.|MailTips\.|SMTP\.|IMAP\.|POP\.)'
    'Calendar'      = '^(Calendars\.|Place\.)'
    'Contacts'      = '^(Contacts\.|People\.)'
    'SharePoint'    = '^(Sites\.|Files\.|AllSites\.|MyFiles\.|TermStore\.|Lists?\.)'
    'Teams'         = '^(Team|Channel|Chat|OnlineMeeting|CallRecords|Calls\.|TeamsApp|TeamsTab|TeamMember)'
    'Reporting'     = '^(Reports\.|AuditLog\.|ActivityFeed\.|ServiceHealth\.|ServiceMessage\.|SecurityEvents\.|ThreatIntelligence\.)'
    'Directory'     = '^(Directory\.|User\.|Group\.|GroupMember\.|Application\.|RoleManagement\.|Organization\.)'
}

# App display name / publisher => app category (report tab). First match wins.
# Override or extend via -AppCategoryMap.
$script:AppCategories = [ordered]@{
    'MailStore'     = 'MailStore'
    'Backup'        = 'Veeam|AvePoint|Acronis|Barracuda|Commvault|Druva|Keepit|Datto|Spanning|Rubrik|Cohesity|Hornetsecurity|Altaro|Dropsuite|Afi\b|NAKIVO|CloudAlly|Arcserve|Synology|Redstor|SkyKick|Backupify|Metallic|Unitrends|Kaseya|Backup|Archiv'
    'Salesforce'    = 'Salesforce|Einstein|Pardot|Slack for Salesforce|Tableau'
    'OutlookOffice' = '^Microsoft Office|^Office 365|^Office|Outlook|Microsoft 365|OneNote|Word|Excel|PowerPoint|Loop'
    'Reporting'     = 'Report|Monitor|Analytics|Insight|Audit|SIEM|Sentinel|Splunk|Power BI|Grafana|PRTG|Quest|ManageEngine|AdminDroid|CoreView|Varonis'
    'Teams'         = 'Teams|Webex|Zoom|Poly|Yealink|Anynode|Ascom|Unify'
    'SharePoint'    = 'SharePoint|OneDrive|ShareGate|Metalogix|Migrat'
    'Exchange'      = 'Exchange|Mail|SMTP|Exclaimer|CodeTwo|Signature|Scan|Printer|Ricoh|Canon|Konica|Xerox|Kyocera|Ticket|Valuemation|ServiceNow|Jira|Zendesk'
}

$script:ReportTabs = @('Exchange', 'SharePoint', 'Teams', 'Reporting', 'MailStore', 'Backup', 'Salesforce', 'OutlookOffice', 'Other')

#endregion

#region ---- private helpers ---------------------------------------------------

function New-AppReportFolder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        $null = New-Item -ItemType Directory -Path $Path -Force
    }
    (Resolve-Path -LiteralPath $Path).ProviderPath
}

function ConvertTo-AppReportDate {
    param($Value)
    if ($null -eq $Value -or '' -eq $Value) { return $null }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime() }
    [datetime]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal)
}

function Get-AppReportProp {
    # StrictMode-safe property access for PSObject / hashtable
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { return $Object[$Name] }
    $p = $Object.PSObject.Properties[$Name]
    if ($p) { $p.Value } else { $null }
}

function Get-GraphCollection {
    param([Parameter(Mandatory)][string]$Uri)
    do {
        $page = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
        foreach ($entry in @(Get-AppReportProp $page 'value')) { if ($null -ne $entry) { $entry } }
        $Uri = Get-AppReportProp $page '@odata.nextLink'
    } while ($Uri)
}

function Resolve-DirectoryObject {
    # Bulk resolve object ids => @{ id = @{ Type; DisplayName; UPN } }
    param([string[]]$Id)
    $map = @{}
    $ids = @($Id | Where-Object { $_ } | Sort-Object -Unique)
    for ($i = 0; $i -lt $ids.Count; $i += 1000) {
        $chunk = $ids[$i..([Math]::Min($i + 999, $ids.Count - 1))]
        $body = @{ ids = @($chunk) } | ConvertTo-Json
        $res = Invoke-MgGraphRequest -Method POST -Uri "$script:GraphBase/directoryObjects/getByIds" -Body $body -ContentType 'application/json' -OutputType PSObject
        foreach ($o in @(Get-AppReportProp $res 'value')) {
            $map[$o.id] = [pscustomobject]@{
                Type        = ((Get-AppReportProp $o '@odata.type') -replace '^#microsoft\.graph\.', '')
                DisplayName = Get-AppReportProp $o 'displayName'
                UPN         = (@((Get-AppReportProp $o 'userPrincipalName'), (Get-AppReportProp $o 'mail')) | Where-Object { $_ } | Select-Object -First 1)
            }
        }
    }
    $map
}

function Get-PermissionCategory {
    param([string]$Permission)
    foreach ($k in $script:PermissionCategories.Keys) {
        if ($Permission -match $script:PermissionCategories[$k]) { return $k }
    }
    'Other'
}

function Get-AppCategory {
    param([string]$Name, [string]$Publisher, [string[]]$PermissionCategory, [System.Collections.IDictionary]$Map)
    foreach ($k in $Map.Keys) {
        if ($Name -match $Map[$k] -or ($Publisher -and $Publisher -match $Map[$k])) {
            if ($k -eq 'OutlookOffice' -and $Name -notmatch $Map[$k]) { continue }
            return $k
        }
    }
    # Fallback: dominant workload of the permissions
    $pc = @($PermissionCategory)
    if ($pc -contains 'EWS' -or $pc -contains 'Mail' -or $pc -contains 'ExchangeAdmin' -or $pc -contains 'Calendar' -or $pc -contains 'Contacts') { return 'Exchange' }
    if ($pc -contains 'SharePoint') { return 'SharePoint' }
    if ($pc -contains 'Teams')      { return 'Teams' }
    if ($pc -contains 'Reporting')  { return 'Reporting' }
    'Other'
}

function Test-ExchangeConnection {
    try { $null = Get-Command Get-OrganizationConfig -ErrorAction Stop; $true } catch { $false }
}

function Assert-GraphConnection {
    $ctx = Get-MgContext
    if (-not $ctx) { throw 'Not connected to Microsoft Graph. Run Connect-AppReport first.' }
    if ($ctx.Environment -and $ctx.Environment -ne 'Global') { throw 'AppReport targets the commercial (Global) Microsoft Graph endpoint.' }
    $ctx
}

#endregion

#region ---- configuration ----------------------------------------------------

function Set-AppReportConfig {
    <#
    .SYNOPSIS
    Saves tenant / app connection settings as JSON. Secrets are stored as DPAPI-protected SecureString.
    .DESCRIPTION
    The JSON file can be used as a splat for Connect-AppReport (-ConfigPath).
    ClientSecret is written with ConvertFrom-SecureString (Windows DPAPI: readable only by the same user on the same machine).
    The target folder is created when missing (default C:\temp\config).
    .EXAMPLE
    Set-AppReportConfig -TenantId contoso.onmicrosoft.com -ClientId 11111111-2222-3333-4444-555555555555 -ClientSecret (Read-Host -AsSecureString 'Secret')
    .EXAMPLE
    Set-AppReportConfig -TenantId contoso.onmicrosoft.com -ClientId $appId -CertificateThumbprint 'A1B2C3...' -Organization contoso.onmicrosoft.com -Path C:\temp\config\contoso.json
    #>
    [CmdletBinding(DefaultParameterSetName = 'Secret')]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory, ParameterSetName = 'Secret')][securestring]$ClientSecret,
        [Parameter(Mandatory, ParameterSetName = 'Certificate')][string]$CertificateThumbprint,
        [string]$Organization,
        [string]$Path = $script:DefaultConfig
    )
    $folder = New-AppReportFolder -Path (Split-Path -Path $Path -Parent)
    $cfg = [ordered]@{
        TenantId     = $TenantId
        ClientId     = $ClientId
        Organization = $(if ($Organization) { $Organization } elseif ($TenantId -notmatch '^[0-9a-fA-F-]{36}$') { $TenantId } else { $null })
        AuthType     = $PSCmdlet.ParameterSetName
        SavedBy      = '{0}\{1}' -f $env:USERDOMAIN, $env:USERNAME
        SavedOn      = $env:COMPUTERNAME
        SavedAt      = (Get-Date).ToString('o')
    }
    if ($PSCmdlet.ParameterSetName -eq 'Secret') {
        if ($PSVersionTable.PSEdition -eq 'Core' -and -not $IsWindows) {
            Write-Warning 'DPAPI is not available on this platform; the secret is only obfuscated, not encrypted.'
        }
        $cfg.ClientSecret = $ClientSecret | ConvertFrom-SecureString
    } else {
        $cfg.CertificateThumbprint = $CertificateThumbprint
    }
    $file = Join-Path $folder (Split-Path -Path $Path -Leaf)
    $cfg | ConvertTo-Json | Set-Content -LiteralPath $file -Encoding UTF8
    Write-Verbose "Config written: $file"
    Get-Item -LiteralPath $file
}

function Get-AppReportConfig {
    <#
    .SYNOPSIS
    Reads an AppReport JSON config and returns a hashtable usable as splat for Connect-AppReport.
    .EXAMPLE
    $cfg = Get-AppReportConfig; Connect-AppReport @cfg
    #>
    [CmdletBinding()]
    param([string]$Path = $script:DefaultConfig)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Config not found: $Path (create it with Set-AppReportConfig)" }
    $json = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    $splat = @{}
    foreach ($k in 'TenantId', 'ClientId', 'CertificateThumbprint', 'Organization') {
        $v = Get-AppReportProp $json $k
        if ($v) { $splat[$k] = [string]$v }
    }
    $sec = Get-AppReportProp $json 'ClientSecret'
    if ($sec) { $splat.ClientSecret = $sec | ConvertTo-SecureString }
    $splat
}

#endregion

#region ---- connection --------------------------------------------------------

function Connect-AppReport {
    <#
    .SYNOPSIS
    Connects to Microsoft Graph (and optionally Exchange Online) interactively, with client secret, certificate or a JSON config.
    .DESCRIPTION
    Interactive   : Connect-MgGraph with delegated read scopes.
    Secret        : app-only, ClientSecretCredential. Exchange Online via OAuth token (-AccessToken, EXO module 3.1+).
    Certificate   : app-only with thumbprint (CurrentUser/LocalMachine store) for Graph and Exchange Online.
    ConfigPath    : JSON created by Set-AppReportConfig (or any JSON with matching keys) used as splat.
    App-only requirements: Graph Application.Read.All + Directory.Read.All (application);
    Exchange: Exchange.ManageAsApp + directory role "Global Reader" or "View-Only Organization Management".
    .EXAMPLE
    Connect-AppReport -TenantId contoso.onmicrosoft.com -Exchange
    .EXAMPLE
    Connect-AppReport -ConfigPath C:\temp\config\AppReport.config.json -Exchange
    .EXAMPLE
    Connect-AppReport -TenantId $tid -ClientId $appId -CertificateThumbprint $thumb -Organization contoso.onmicrosoft.com -Exchange
    #>
    [CmdletBinding(DefaultParameterSetName = 'Interactive')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Interactive')]
        [Parameter(Mandatory, ParameterSetName = 'Secret')]
        [Parameter(Mandatory, ParameterSetName = 'Certificate')]
        [string]$TenantId,

        [Parameter(Mandatory, ParameterSetName = 'Secret')]
        [Parameter(Mandatory, ParameterSetName = 'Certificate')]
        [string]$ClientId,

        [Parameter(Mandatory, ParameterSetName = 'Secret')][securestring]$ClientSecret,
        [Parameter(Mandatory, ParameterSetName = 'Certificate')][string]$CertificateThumbprint,

        [Parameter(Mandatory, ParameterSetName = 'Config')][string]$ConfigPath,

        # Exchange organization (contoso.onmicrosoft.com) – required for app-only EXO
        [string]$Organization,
        # Also connect Exchange Online (needed for access policies, RBAC, EWS config)
        [switch]$Exchange,
        # Delegated: UPN for Connect-ExchangeOnline
        [string]$UserPrincipalName
    )
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    if ($PSCmdlet.ParameterSetName -eq 'Config') {
        $splat = Get-AppReportConfig -Path $ConfigPath
        if ($Organization) { $splat.Organization = $Organization }
        if ($Exchange)     { $splat.Exchange = $true }
        if (-not $splat.ContainsKey('ClientSecret') -and -not $splat.ContainsKey('CertificateThumbprint')) { $null = $splat.Remove('ClientId') }
        return Connect-AppReport @splat
    }

    switch ($PSCmdlet.ParameterSetName) {
        'Interactive' {
            Connect-MgGraph -TenantId $TenantId -Scopes $script:GraphScopes -ContextScope Process -NoWelcome
        }
        'Secret' {
            $cred = New-Object System.Management.Automation.PSCredential ($ClientId, $ClientSecret)
            Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $cred -ContextScope Process -NoWelcome
        }
        'Certificate' {
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -ContextScope Process -NoWelcome
        }
    }
    $ctx = Assert-GraphConnection
    Write-Verbose ("Graph connected: tenant {0}, auth {1}" -f $ctx.TenantId, $ctx.AuthType)

    if ($Exchange) {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        if (-not $Organization -and $TenantId -notmatch '^[0-9a-fA-F-]{36}$') { $Organization = $TenantId }
        switch ($PSCmdlet.ParameterSetName) {
            'Interactive' {
                $p = @{ ShowBanner = $false }
                if ($UserPrincipalName) { $p.UserPrincipalName = $UserPrincipalName }
                Connect-ExchangeOnline @p
            }
            'Certificate' {
                if (-not $Organization) { throw 'Organization (xxx.onmicrosoft.com) is required for app-only Exchange Online.' }
                Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -Organization $Organization -ShowBanner:$false
            }
            'Secret' {
                if (-not $Organization) { throw 'Organization (xxx.onmicrosoft.com) is required for app-only Exchange Online.' }
                $bstr  = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret)
                $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
                [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
                $body = @{
                    client_id     = $ClientId
                    client_secret = $plain
                    scope         = 'https://outlook.office365.com/.default'
                    grant_type    = 'client_credentials'
                }
                $tok = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $body -ContentType 'application/x-www-form-urlencoded'
                $plain = $null; $body = $null
                Connect-ExchangeOnline -AccessToken $tok.access_token -Organization $Organization -ShowBanner:$false
            }
        }
    }
    $ctx
}

#endregion

#region ---- Graph inventory ---------------------------------------------------

function Get-AppReportInventory {
    <#
    .SYNOPSIS
    Graph inventory of app registrations / enterprise apps with permissions on M365 workload APIs.
    .DESCRIPTION
    Returns an object with:
      Permissions  – one row per app/permission/evidence (Configured = manifest, Granted = consent / app role assignment)
      Credentials  – one row per secret / certificate (app registration and service principal) incl. expiry state
      Apps         – one row per app (aggregated, category = report tab, NextExpiry, ScopeType)
    .PARAMETER AllPermissions
    Include every permission on the evaluated resource APIs (default: only workload categories, no plain Directory/User.Read).
    .PARAMETER ExpiringDays
    Credentials ending within this many days are marked 'Expiring'. Default 30.
    .EXAMPLE
    $inv = Get-AppReportInventory
    $inv.Apps | Sort-Object AppCategory | Format-Table AppName,AppCategory,PermissionCategories,NextExpiry,CredentialState
    #>
    [CmdletBinding()]
    param(
        [switch]$AllPermissions,
        [switch]$ExcludeMicrosoftFirstParty,
        [int]$ExpiringDays = 30,
        [System.Collections.IDictionary]$AppCategoryMap,
        [switch]$SkipUserResolution
    )
    $ctx = Assert-GraphConnection
    $map = [ordered]@{}
    if ($AppCategoryMap) { foreach ($k in $AppCategoryMap.Keys) { $map[$k] = $AppCategoryMap[$k] } }
    foreach ($k in $script:AppCategories.Keys) { if (-not $map.Contains($k)) { $map[$k] = $script:AppCategories[$k] } }
    $now = (Get-Date).ToUniversalTime()

    Write-Verbose 'Reading applications'
    $apps = @(Get-GraphCollection "$script:GraphBase/applications?`$select=id,appId,displayName,createdDateTime,signInAudience,publisherDomain,requiredResourceAccess,passwordCredentials,keyCredentials,notes&`$top=999")
    Write-Verbose 'Reading service principals'
    $sps = @(Get-GraphCollection "$script:GraphBase/servicePrincipals?`$select=id,appId,displayName,servicePrincipalType,accountEnabled,appOwnerOrganizationId,publisherName,verifiedPublisher,appRoles,oauth2PermissionScopes,passwordCredentials,keyCredentials,tags&`$top=999")

    $spById = @{}; $spByApp = @{}; $appByApp = @{}
    foreach ($sp in $sps)   { $spById[$sp.id] = $sp; $spByApp[$sp.appId] = $sp }
    foreach ($a in $apps)   { $appByApp[$a.appId] = $a }

    $resources = @($sps | Where-Object { $script:ResourceApis.Contains($_.appId) })
    $catalog = @{}
    foreach ($r in $resources) {
        foreach ($role in @(Get-AppReportProp $r 'appRoles'))               { if ($role) { $catalog['{0}|Role|{1}'  -f $r.appId, $role.id]  = $role.value } }
        foreach ($sc   in @(Get-AppReportProp $r 'oauth2PermissionScopes')) { if ($sc)   { $catalog['{0}|Scope|{1}' -f $r.appId, $sc.id]    = $sc.value } }
    }

    $perm = New-Object 'System.Collections.Generic.List[object]'
    $unresolved = 0

    $addPerm = {
        param($ClientSp, $LocalApp, $Resource, [string]$Value, [string]$Type, [string]$Evidence, [string]$ConsentType, [string]$PrincipalId)
        $cat = Get-PermissionCategory $Value
        if (-not $AllPermissions -and $cat -in @('Other', 'Directory')) { return }
        $name = $(if ($LocalApp) { $LocalApp.displayName } elseif ($ClientSp) { $ClientSp.displayName } else { '' })
        $appId = $(if ($LocalApp) { $LocalApp.appId } else { $ClientSp.appId })
        $owner = Get-AppReportProp $ClientSp 'appOwnerOrganizationId'
        $firstParty = [bool]($owner -and $owner -in $script:MicrosoftOwnerTenants)
        if ($ExcludeMicrosoftFirstParty -and $firstParty) { return }
        $scopeType = switch ($Type) {
            'Application' {
                if ($Value -match '\.Selected$') { 'Selected resources' }
                elseif ($cat -in 'Mail', 'Calendar', 'Contacts', 'EWS') { 'Tenant-wide (all mailboxes)' }
                elseif ($cat -eq 'ExchangeAdmin') { 'Exchange admin (role-based)' }
                else { 'Tenant-wide' }
            }
            default {
                if ($ConsentType -eq 'Principal') { 'Single user (delegated)' }
                elseif ($ConsentType -eq 'AllPrincipals') { 'All users (delegated, signed-in context)' }
                else { 'Delegated (requested)' }
            }
        }
        $perm.Add([pscustomobject][ordered]@{
            TenantId                 = $ctx.TenantId
            AppName                  = $name
            AppId                    = $appId
            AppRegistrationObjectId  = Get-AppReportProp $LocalApp 'id'
            ServicePrincipalObjectId = Get-AppReportProp $ClientSp 'id'
            ServicePrincipalType     = Get-AppReportProp $ClientSp 'servicePrincipalType'
            AccountEnabled           = Get-AppReportProp $ClientSp 'accountEnabled'
            LocalAppRegistration     = [bool]$LocalApp
            MicrosoftFirstParty      = $firstParty
            Publisher                = Get-AppReportProp $ClientSp 'publisherName'
            Resource                 = $script:ResourceApis[$Resource.appId]
            ResourceAppId            = $Resource.appId
            Permission               = $Value
            PermissionCategory       = $cat
            PermissionType           = $Type
            Evidence                 = $Evidence
            ConsentType              = $ConsentType
            ConsentUserObjectId      = $PrincipalId
            ConsentUser              = $null
            ScopeType                = $scopeType
            UsesEWS                  = ($cat -eq 'EWS')
        })
    }

    # 1. requested in local manifests
    foreach ($a in $apps) {
        foreach ($req in @(Get-AppReportProp $a 'requiredResourceAccess')) {
            if (-not $req -or -not $script:ResourceApis.Contains($req.resourceAppId)) { continue }
            foreach ($acc in @($req.resourceAccess)) {
                $v = $catalog['{0}|{1}|{2}' -f $req.resourceAppId, $acc.type, $acc.id]
                if (-not $v) { $unresolved++; continue }
                $t = $(if ($acc.type -eq 'Role') { 'Application' } else { 'Delegated' })
                $res = $spByApp[$req.resourceAppId]
                if (-not $res) { $res = [pscustomobject]@{ appId = $req.resourceAppId } }
                & $addPerm $spByApp[$a.appId] $a $res $v $t 'Configured' '' ''
            }
        }
    }
    # 2. application permission grants
    foreach ($r in $resources) {
        foreach ($as in @(Get-GraphCollection "$script:GraphBase/servicePrincipals/$($r.id)/appRoleAssignedTo?`$top=999")) {
            if ($as.principalType -ne 'ServicePrincipal') { continue }
            $v = $catalog['{0}|Role|{1}' -f $r.appId, $as.appRoleId]
            if (-not $v) { $unresolved++; continue }
            $client = $spById[$as.principalId]
            if (-not $client) { Write-Warning "Client service principal missing: $($as.principalId)"; continue }
            & $addPerm $client $appByApp[$client.appId] $r $v 'Application' 'Granted' 'Application' ''
        }
    }
    # 3. delegated grants
    foreach ($g in @(Get-GraphCollection "$script:GraphBase/oauth2PermissionGrants")) {
        $r = $spById[$g.resourceId]
        if (-not $r -or -not $script:ResourceApis.Contains($r.appId)) { continue }
        $client = $spById[$g.clientId]
        if (-not $client) { Write-Warning "Client service principal missing: $($g.clientId)"; continue }
        foreach ($v in (([string]$g.scope).Trim() -split '\s+')) {
            if ($v) { & $addPerm $client $appByApp[$client.appId] $r $v 'Delegated' 'Granted' $g.consentType (Get-AppReportProp $g 'principalId') }
        }
    }
    if ($unresolved) { Write-Warning "$unresolved permission IDs could not be resolved." }

    if (-not $SkipUserResolution) {
        $uids = @($perm | Where-Object ConsentUserObjectId | Select-Object -ExpandProperty ConsentUserObjectId -Unique)
        if ($uids.Count) {
            $um = Resolve-DirectoryObject -Id $uids
            foreach ($p in $perm) { if ($p.ConsentUserObjectId -and $um[$p.ConsentUserObjectId]) { $p.ConsentUser = $um[$p.ConsentUserObjectId].UPN } }
        }
    }

    # Credentials (only apps that appear in the permission inventory, plus all local apps with -AllPermissions)
    $appIds = @($perm | Select-Object -ExpandProperty AppId -Unique)
    $cred = New-Object 'System.Collections.Generic.List[object]'
    $addCred = {
        param($Owner, [string]$OwnerType, $C, [string]$Kind)
        $end = ConvertTo-AppReportDate (Get-AppReportProp $C 'endDateTime')
        $days = $(if ($end) { [int][Math]::Floor(($end - $now).TotalDays) } else { $null })
        $state = $(if (-not $end) { 'Unknown' } elseif ($end -lt $now) { 'Expired' } elseif ($days -le $ExpiringDays) { 'Expiring' } else { 'Valid' })
        $cred.Add([pscustomobject][ordered]@{
            AppName        = $Owner.displayName
            AppId          = $Owner.appId
            OwnerType      = $OwnerType
            CredentialType = $Kind
            DisplayName    = Get-AppReportProp $C 'displayName'
            KeyId          = Get-AppReportProp $C 'keyId'
            Hint           = Get-AppReportProp $C 'hint'
            Usage          = Get-AppReportProp $C 'usage'
            StartDate      = ConvertTo-AppReportDate (Get-AppReportProp $C 'startDateTime')
            EndDate        = $end
            DaysLeft       = $days
            State          = $state
        })
    }
    foreach ($id in $appIds) {
        $a = $appByApp[$id]; $sp = $spByApp[$id]
        if ($a)  { foreach ($c in @(Get-AppReportProp $a 'passwordCredentials')) { if ($c) { & $addCred $a 'AppRegistration' $c 'Secret' } }
                   foreach ($c in @(Get-AppReportProp $a 'keyCredentials'))      { if ($c) { & $addCred $a 'AppRegistration' $c 'Certificate' } } }
        if ($sp) { foreach ($c in @(Get-AppReportProp $sp 'passwordCredentials')) { if ($c) { & $addCred $sp 'ServicePrincipal' $c 'Secret' } }
                   foreach ($c in @(Get-AppReportProp $sp 'keyCredentials'))      { if ($c -and (Get-AppReportProp $c 'usage') -eq 'Verify') { & $addCred $sp 'ServicePrincipal' $c 'Certificate' } } }
    }

    # App aggregation
    $appRows = foreach ($grp in ($perm | Group-Object AppId)) {
        $first = $grp.Group[0]
        $pcats = @($grp.Group.PermissionCategory | Sort-Object -Unique)
        $creds = @($cred | Where-Object AppId -eq $grp.Name)
        $valid = @($creds | Where-Object { $_.State -in 'Valid', 'Expiring' } | Sort-Object EndDate)
        $credState = $(if (-not $creds.Count) { $(if ($first.LocalAppRegistration) { 'NoCredential' } else { 'External/Managed' }) }
                       elseif (-not $valid.Count) { 'AllExpired' }
                       elseif ($valid[0].State -eq 'Expiring') { 'Expiring' }
                       elseif (@($creds | Where-Object State -eq 'Expired').Count) { 'PartlyExpired' }
                       else { 'Valid' })
        [pscustomobject][ordered]@{
            AppName              = $first.AppName
            AppId                = $grp.Name
            AppCategory          = Get-AppCategory -Name $first.AppName -Publisher $first.Publisher -PermissionCategory $pcats -Map $map
            Publisher            = $first.Publisher
            ServicePrincipalType = $first.ServicePrincipalType
            LocalAppRegistration = $first.LocalAppRegistration
            MicrosoftFirstParty  = $first.MicrosoftFirstParty
            AccountEnabled       = $first.AccountEnabled
            PermissionCategories = ($pcats -join ', ')
            ApplicationPerms     = ((@($grp.Group | Where-Object PermissionType -eq 'Application' | Select-Object -ExpandProperty Permission -Unique)) -join ', ')
            DelegatedPerms       = ((@($grp.Group | Where-Object PermissionType -eq 'Delegated'   | Select-Object -ExpandProperty Permission -Unique)) -join ', ')
            GrantedApplication   = [bool](@($grp.Group | Where-Object { $_.Evidence -eq 'Granted' -and $_.PermissionType -eq 'Application' }).Count)
            UsesEWS              = [bool](@($grp.Group | Where-Object UsesEWS).Count)
            ScopeType            = ((@($grp.Group.ScopeType | Sort-Object -Unique)) -join '; ')
            CredentialCount      = $creds.Count
            NextExpiry           = $(if ($valid.Count) { $valid[0].EndDate } else { $null })
            LastExpired          = ($creds | Where-Object State -eq 'Expired' | Sort-Object EndDate -Descending | Select-Object -First 1 -ExpandProperty EndDate)
            CredentialState      = $credState
            AccessPolicy         = $null
            ExchangeRbacScope    = $null
            EffectiveMailScope   = $null
        }
    }

    [pscustomobject]@{
        TenantId    = $ctx.TenantId
        Permissions = @($perm | Sort-Object AppName, Resource, PermissionType, Permission, Evidence)
        Credentials = @($cred | Sort-Object EndDate)
        Apps        = @($appRows | Sort-Object AppCategory, AppName)
    }
}

#endregion

#region ---- Exchange ----------------------------------------------------------

function Get-AppReportAccessPolicy {
    <#
    .SYNOPSIS
    Inventory of Exchange Online Application Access Policies (legacy app scoping) with resolved scope group / user.
    .PARAMETER ExpandMembers
    Adds one row per member of the scope group (recursive for nested distribution / mail-enabled security groups).
    .EXAMPLE
    Get-AppReportAccessPolicy -ExpandMembers | Group-Object ScopeName
    #>
    [CmdletBinding()]
    param([switch]$ExpandMembers)
    if (-not (Test-ExchangeConnection)) { throw 'Not connected to Exchange Online (Connect-AppReport -Exchange).' }
    foreach ($p in @(Get-ApplicationAccessPolicy -ErrorAction Stop)) {
        $scopeId = [string]$p.ScopeIdentity
        $recip = $null
        $raw = Get-AppReportProp $p 'ScopeIdentityRaw'
        try { $recip = Get-Recipient -Identity ($(if ($raw) { [string]$raw } else { $scopeId })) -ErrorAction Stop } catch { }
        $base = [ordered]@{
            AppId          = [string]$p.AppId
            AccessRight    = [string]$p.AccessRight
            ScopeIdentity  = $scopeId
            ScopeName      = $(if ($recip) { $recip.DisplayName } else { $scopeId })
            ScopeSmtp      = $(if ($recip) { [string]$recip.PrimarySmtpAddress } else { $null })
            ScopeType      = $(if ($recip) { [string]$recip.RecipientTypeDetails } else { 'Unresolved' })
            Description    = [string]$p.Description
            IsValid        = Get-AppReportProp $p 'IsValid'
            Member         = $null
            MemberSmtp     = $null
            MemberType     = $null
        }
        if ($ExpandMembers -and $recip -and $recip.RecipientTypeDetails -match 'Group') {
            $seen = @{}
            $queue = New-Object System.Collections.Queue
            $queue.Enqueue([string]$recip.Guid)
            while ($queue.Count) {
                $g = $queue.Dequeue()
                if ($seen[$g]) { continue }; $seen[$g] = $true
                foreach ($m in @(Get-DistributionGroupMember -Identity $g -ResultSize Unlimited -ErrorAction SilentlyContinue)) {
                    if ($m.RecipientTypeDetails -match 'Group') { $queue.Enqueue([string]$m.Guid); continue }
                    $row = [pscustomobject]$base
                    $row.Member = $m.DisplayName; $row.MemberSmtp = [string]$m.PrimarySmtpAddress; $row.MemberType = [string]$m.RecipientTypeDetails
                    $row
                }
            }
        } else {
            [pscustomobject]$base
        }
    }
}

function Get-AppReportExchangeRbac {
    <#
    .SYNOPSIS
    Inventory of Exchange Online RBAC for Applications (Application Mail.Read, Application EWS.AccessAsApp, ...) with scope.
    .DESCRIPTION
    Reads Get-ServicePrincipal and Get-ManagementRoleAssignment (RoleAssigneeType ServicePrincipal) and resolves
    the CustomResourceScope as management scope (RecipientRestrictionFilter) or administrative unit.
    .EXAMPLE
    Get-AppReportExchangeRbac | Format-Table AppName,Role,ScopeType,ScopeName,ScopeFilter
    #>
    [CmdletBinding()]
    param()
    if (-not (Test-ExchangeConnection)) { throw 'Not connected to Exchange Online (Connect-AppReport -Exchange).' }
    $spMap = @{}
    foreach ($s in @(Get-ServicePrincipal -ErrorAction Stop)) {
        foreach ($key in @([string]$s.Identity, [string]$s.ObjectId, [string]$s.Name, [string]$s.DisplayName)) { if ($key) { $spMap[$key] = $s } }
    }
    $scopes = @{}
    foreach ($m in @(Get-ManagementScope -ErrorAction SilentlyContinue)) { $scopes[[string]$m.Name] = $m; $scopes[[string]$m.Identity] = $m }

    $assign = @(Get-ManagementRoleAssignment -ErrorAction Stop | Where-Object { [string]$_.RoleAssigneeType -eq 'ServicePrincipal' })
    foreach ($a in $assign) {
        $sp = $spMap[[string]$a.RoleAssignee]; if (-not $sp) { $sp = $spMap[[string]$a.RoleAssigneeName] }
        $scopeName = [string](Get-AppReportProp $a 'CustomResourceScope')
        $scopeType = 'Tenant-wide'; $filter = $null
        if ($scopeName) {
            $ms = $scopes[$scopeName]
            if ($ms) { $scopeType = 'ManagementScope'; $filter = [string]$ms.RecipientFilter }
            else      { $scopeType = 'AdministrativeUnit' }
        }
        [pscustomobject][ordered]@{
            AppName          = $(if ($sp) { $sp.DisplayName } else { [string]$a.RoleAssigneeName })
            AppId            = $(if ($sp) { [string]$sp.AppId } else { $null })
            ServicePrincipal = $(if ($sp) { [string]$sp.ObjectId } else { [string]$a.RoleAssignee })
            Role             = [string]$a.Role
            AssignmentName   = [string]$a.Name
            Enabled          = $a.Enabled
            ScopeType        = $scopeType
            ScopeName        = $scopeName
            ScopeFilter      = $filter
            WriteScope       = [string]$a.RecipientWriteScope
        }
    }
}

function Get-AppReportEwsConfig {
    <#
    .SYNOPSIS
    EWS configuration: organization level (all Ews* parameters of Get-OrganizationConfig) and mailbox-level exceptions (CASMailbox).
    .DESCRIPTION
    Rows with Level 'Organization' list each Ews* org parameter.
    Rows with Level 'Mailbox' list mailboxes whose EwsEnabled / EwsApplicationAccessPolicy / EwsAllowList / EwsBlockList
    differ from the organization (= exceptions).
    .EXAMPLE
    Get-AppReportEwsConfig | Where-Object Level -eq 'Mailbox' | Format-Table Identity,Setting,Value,OrgValue
    #>
    [CmdletBinding()]
    param([switch]$SkipMailboxes)
    if (-not (Test-ExchangeConnection)) { throw 'Not connected to Exchange Online (Connect-AppReport -Exchange).' }
    $org = Get-OrganizationConfig
    $orgVals = @{}
    foreach ($p in ($org.PSObject.Properties | Where-Object { $_.Name -like 'Ews*' -or $_.Name -eq 'OAuth2ClientProfileEnabled' } | Sort-Object Name)) {
        $v = $(if ($p.Value -is [System.Collections.IEnumerable] -and $p.Value -isnot [string]) { (@($p.Value) -join '; ') } else { [string]$p.Value })
        $orgVals[$p.Name] = $v
        [pscustomobject][ordered]@{ Level = 'Organization'; Identity = [string]$org.Name; PrimarySmtpAddress = $null; Setting = $p.Name; Value = $v; OrgValue = $v; Exception = $false }
    }
    if ($SkipMailboxes) { return }

    $props = 'EwsEnabled', 'EwsApplicationAccessPolicy', 'EwsAllowList', 'EwsBlockList', 'EwsAllowOutlook', 'EwsAllowMacOutlook'
    $cas = $(if (Get-Command Get-EXOCASMailbox -ErrorAction SilentlyContinue) {
                Get-EXOCASMailbox -ResultSize Unlimited -PropertySets Minimum, Ews -ErrorAction Stop
             } else { Get-CASMailbox -ResultSize Unlimited -ErrorAction Stop })
    $orgEws = $orgVals['EwsEnabled']
    foreach ($m in $cas) {
        foreach ($s in $props) {
            $pp = $m.PSObject.Properties[$s]; if (-not $pp) { continue }
            $v = $(if ($pp.Value -is [System.Collections.IEnumerable] -and $pp.Value -isnot [string]) { (@($pp.Value) -join '; ') } else { [string]$pp.Value })
            if ('' -eq $v) { continue }
            $ov = $orgVals[$s]
            $isExc = $(switch ($s) {
                'EwsEnabled' { ($v -eq 'False' -and $orgEws -ne 'False') -or ($v -eq 'True' -and $orgEws -eq 'False') }
                default      { $v -ne $ov }
            })
            if ($isExc) {
                [pscustomobject][ordered]@{ Level = 'Mailbox'; Identity = [string]$m.Identity; PrimarySmtpAddress = [string]$m.PrimarySmtpAddress; Setting = $s; Value = $v; OrgValue = $ov; Exception = $true }
            }
        }
    }
}

#endregion

#region ---- scope correlation -------------------------------------------------

function Get-AppReportScope {
    <#
    .SYNOPSIS
    Correlates Graph permissions, Application Access Policies and Exchange RBAC into one row per app and target scope,
    grouped by target (user / group / management scope / tenant).
    .EXAMPLE
    $scope = Get-AppReportScope -Inventory $inv -AccessPolicy $aap -ExchangeRbac $rbac
    $scope | Group-Object Target | Sort-Object Count -Descending
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Inventory,
        $AccessPolicy,
        $ExchangeRbac
    )
    $aapByApp = @{}; foreach ($p in @($AccessPolicy)) { if ($p) { if (-not $aapByApp[$p.AppId]) { $aapByApp[$p.AppId] = @() }; $aapByApp[$p.AppId] += $p } }
    $rbByApp  = @{}; foreach ($r in @($ExchangeRbac)) { if ($r -and $r.AppId) { if (-not $rbByApp[$r.AppId]) { $rbByApp[$r.AppId] = @() }; $rbByApp[$r.AppId] += $r } }
    $appInfo = @{}; foreach ($a in $Inventory.Apps) { $appInfo[$a.AppId] = $a }

    $rows = New-Object 'System.Collections.Generic.List[object]'
    $mailCats = 'Mail', 'Calendar', 'Contacts', 'EWS'

    foreach ($grp in ($Inventory.Permissions | Where-Object { $_.Evidence -eq 'Granted' } | Group-Object AppId)) {
        $appId = $grp.Name; $info = $appInfo[$appId]
        $appPerms = @($grp.Group | Where-Object PermissionType -eq 'Application')
        $mailApp  = @($appPerms | Where-Object { $_.PermissionCategory -in $mailCats })
        $other    = @($appPerms | Where-Object { $_.PermissionCategory -notin $mailCats })
        $del      = @($grp.Group | Where-Object PermissionType -eq 'Delegated')

        if ($mailApp.Count) {
            $pol = @($aapByApp[$appId] | Where-Object { $_ })
            $permText = (@($mailApp.Permission | Sort-Object -Unique) -join ', ')
            if ($pol.Count) {
                foreach ($p in ($pol | Sort-Object AccessRight, ScopeName, ScopeSmtp, MemberSmtp -Unique)) {
                    $rows.Add([pscustomobject][ordered]@{
                        Target = $(if ($p.MemberSmtp) { $p.MemberSmtp } elseif ($p.ScopeSmtp) { $p.ScopeSmtp } else { $p.ScopeName })
                        TargetType = $(if ($p.Member) { 'User (via group)' } else { $p.ScopeType })
                        ViaGroup = $(if ($p.Member) { $p.ScopeName } else { $null })
                        AppName = $info.AppName; AppId = $appId; AppCategory = $info.AppCategory
                        Source = 'Graph app permission + ApplicationAccessPolicy'
                        Restriction = $p.AccessRight; Permissions = $permText; CredentialState = $info.CredentialState
                    })
                }
                $info.AccessPolicy = (@($pol | ForEach-Object { '{0}:{1}' -f $_.AccessRight, $_.ScopeName } | Sort-Object -Unique) -join '; ')
                $info.EffectiveMailScope = $(if (@($pol | Where-Object AccessRight -eq 'RestrictAccess').Count) { 'Restricted (ApplicationAccessPolicy)' } else { 'Tenant-wide minus denied' })
            } else {
                $rows.Add([pscustomobject][ordered]@{
                    Target = 'ALL MAILBOXES'; TargetType = 'Tenant'; ViaGroup = $null
                    AppName = $info.AppName; AppId = $appId; AppCategory = $info.AppCategory
                    Source = 'Graph app permission (no ApplicationAccessPolicy)'
                    Restriction = 'None'; Permissions = $permText; CredentialState = $info.CredentialState
                })
                $info.EffectiveMailScope = 'Tenant-wide (unrestricted)'
            }
        }
        if ($other.Count) {
            $rows.Add([pscustomobject][ordered]@{
                Target = $(if (@($other | Where-Object { $_.Permission -match '\.Selected$' }).Count) { 'SELECTED RESOURCES' } else { 'TENANT' })
                TargetType = 'Tenant'; ViaGroup = $null
                AppName = $info.AppName; AppId = $appId; AppCategory = $info.AppCategory
                Source = 'Graph app permission'
                Restriction = 'None'; Permissions = (@($other.Permission | Sort-Object -Unique) -join ', '); CredentialState = $info.CredentialState
            })
        }
        foreach ($d in ($del | Group-Object { if ($_.ConsentType -eq 'Principal') { $_.ConsentUser } else { 'ALL USERS (delegated)' } })) {
            $rows.Add([pscustomobject][ordered]@{
                Target = $(if ($d.Name) { $d.Name } else { $d.Group[0].ConsentUserObjectId })
                TargetType = $(if ($d.Group[0].ConsentType -eq 'Principal') { 'User (delegated consent)' } else { 'AllPrincipals' })
                ViaGroup = $null
                AppName = $info.AppName; AppId = $appId; AppCategory = $info.AppCategory
                Source = 'Delegated grant'; Restriction = 'Signed-in user permissions'
                Permissions = (@($d.Group.Permission | Sort-Object -Unique) -join ', '); CredentialState = $info.CredentialState
            })
        }
    }
    foreach ($r in @($ExchangeRbac)) {
        if (-not $r) { continue }
        $info = $(if ($r.AppId) { $appInfo[$r.AppId] } else { $null })
        $rows.Add([pscustomobject][ordered]@{
            Target = $(if ($r.ScopeName) { $r.ScopeName } else { 'ALL MAILBOXES' })
            TargetType = $r.ScopeType; ViaGroup = $null
            AppName = $r.AppName; AppId = $r.AppId; AppCategory = $(if ($info) { $info.AppCategory } else { 'Exchange' })
            Source = 'Exchange RBAC for Applications'; Restriction = $(if ($r.ScopeFilter) { $r.ScopeFilter } else { $r.ScopeType })
            Permissions = $r.Role; CredentialState = $(if ($info) { $info.CredentialState } else { $null })
        })
        if ($info) { $info.ExchangeRbacScope = (@(@($info.ExchangeRbacScope) + ('{0} [{1}]' -f $r.Role, $(if ($r.ScopeName) { $r.ScopeName } else { 'Tenant' }))) | Where-Object { $_ }) -join '; ' }
    }
    @($rows | Sort-Object Target, AppName)
}

#endregion

#region ---- orchestration / export --------------------------------------------

function Invoke-AppReport {
    <#
    .SYNOPSIS
    Runs the complete report and writes CSV files (always) and optionally one XLSX workbook with tabs (ImportExcel).
    .DESCRIPTION
    Uses existing Graph / Exchange connections unless connection parameters or -ConfigPath are given.
    Exchange parts (access policies, RBAC, EWS) run when an Exchange Online session exists or -Exchange is set.
    Output files: <OutputPath>\AppReport-<tenant>-<timestamp>-<Sheet>.csv and .xlsx.
    .EXAMPLE
    Invoke-AppReport -TenantId contoso.onmicrosoft.com -Exchange -Xlsx
    .EXAMPLE
    Invoke-AppReport -ConfigPath C:\temp\config\AppReport.config.json -Exchange -Xlsx -ExpandAccessPolicyMembers -OutputPath C:\temp\AppReport
    #>
    [CmdletBinding(DefaultParameterSetName = 'Existing')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Interactive')]
        [Parameter(Mandatory, ParameterSetName = 'Secret')]
        [Parameter(Mandatory, ParameterSetName = 'Certificate')]
        [string]$TenantId,
        [Parameter(Mandatory, ParameterSetName = 'Secret')]
        [Parameter(Mandatory, ParameterSetName = 'Certificate')]
        [string]$ClientId,
        [Parameter(Mandatory, ParameterSetName = 'Secret')][securestring]$ClientSecret,
        [Parameter(Mandatory, ParameterSetName = 'Certificate')][string]$CertificateThumbprint,
        [Parameter(Mandatory, ParameterSetName = 'Config')][string]$ConfigPath,
        [string]$Organization,
        [switch]$Exchange,

        [string]$OutputPath = 'C:\temp\AppReport',
        [switch]$Xlsx,
        [switch]$NoCsv,
        [switch]$AllPermissions,
        [switch]$ExcludeMicrosoftFirstParty,
        [int]$ExpiringDays = 30,
        [switch]$ExpandAccessPolicyMembers,
        [switch]$SkipExchangeRbac,
        [switch]$SkipEws,
        [switch]$SkipEwsMailboxes,
        [System.Collections.IDictionary]$AppCategoryMap
    )
    if ($PSCmdlet.ParameterSetName -ne 'Existing') {
        $c = @{}
        foreach ($k in 'TenantId', 'ClientId', 'ClientSecret', 'CertificateThumbprint', 'ConfigPath', 'Organization', 'Exchange') {
            if ($PSBoundParameters.ContainsKey($k)) { $c[$k] = $PSBoundParameters[$k] }
        }
        $null = Connect-AppReport @c
    }
    $ctx = Assert-GraphConnection

    $inv = Get-AppReportInventory -AllPermissions:$AllPermissions -ExcludeMicrosoftFirstParty:$ExcludeMicrosoftFirstParty -ExpiringDays $ExpiringDays -AppCategoryMap $AppCategoryMap

    $aap = @(); $rbac = @(); $ews = @()
    if (Test-ExchangeConnection) {
        Write-Verbose 'Exchange: application access policies'
        $aap = @(Get-AppReportAccessPolicy -ExpandMembers:$ExpandAccessPolicyMembers)
        if (-not $SkipExchangeRbac) { Write-Verbose 'Exchange: RBAC for applications'; $rbac = @(Get-AppReportExchangeRbac) }
        if (-not $SkipEws)          { Write-Verbose 'Exchange: EWS config';           $ews  = @(Get-AppReportEwsConfig -SkipMailboxes:$SkipEwsMailboxes) }
    } else {
        Write-Warning 'No Exchange Online session – ApplicationAccessPolicy, Exchange RBAC and EWS config are not included (use -Exchange).'
    }
    $scope = Get-AppReportScope -Inventory $inv -AccessPolicy $aap -ExchangeRbac $rbac

    $sheets = [ordered]@{
        Summary          = @($inv.Apps)
        Permissions      = @($inv.Permissions)
        Credentials      = @($inv.Credentials)
        Expired          = @($inv.Credentials | Where-Object State -eq 'Expired')
        Expiring         = @($inv.Credentials | Where-Object State -eq 'Expiring')
    }
    foreach ($t in $script:ReportTabs) { $sheets[$t] = @($inv.Apps | Where-Object AppCategory -eq $t) }
    $sheets['ScopeByTarget']  = @($scope)
    $sheets['AccessPolicies'] = @($aap)
    $sheets['ExchangeRBAC']   = @($rbac)
    $sheets['EWSConfig']      = @($ews)

    $folder = New-AppReportFolder -Path $OutputPath
    $stamp  = Get-Date -Format 'yyyyMMdd-HHmmss'
    $prefix = Join-Path $folder ('AppReport-{0}-{1}' -f $ctx.TenantId, $stamp)
    $files  = @()

    if (-not $NoCsv) {
        foreach ($k in $sheets.Keys) {
            if (-not $sheets[$k].Count) { continue }
            $f = '{0}-{1}.csv' -f $prefix, $k
            $sheets[$k] | Export-Csv -LiteralPath $f -NoTypeInformation -Encoding UTF8 -Delimiter ';'
            $files += $f
        }
    }
    if ($Xlsx) {
        if (-not (Get-Module -ListAvailable ImportExcel)) {
            Write-Warning 'ImportExcel not installed (Install-Module ImportExcel -Scope CurrentUser) – XLSX skipped.'
        } else {
            Import-Module ImportExcel -ErrorAction Stop
            $x = '{0}.xlsx' -f $prefix
            foreach ($k in $sheets.Keys) {
                if (-not $sheets[$k].Count) { continue }
                $p = @{ Path = $x; WorksheetName = $k; TableName = ('t' + $k); TableStyle = 'Medium2'; AutoSize = $true; FreezeTopRow = $true; BoldTopRow = $true }
                if ($k -in 'Expired', 'Expiring', 'Credentials') {
                    $p.ConditionalText = @(
                        (New-ConditionalText -Text 'Expired'  -BackgroundColor '#F8CBAD' -ConditionalTextColor '#9C0006'),
                        (New-ConditionalText -Text 'Expiring' -BackgroundColor '#FFEB9C' -ConditionalTextColor '#9C5700'))
                }
                if ($k -eq 'Summary') {
                    $p.ConditionalText = @(
                        (New-ConditionalText -Text 'AllExpired' -BackgroundColor '#F8CBAD' -ConditionalTextColor '#9C0006'),
                        (New-ConditionalText -Text 'Expiring'   -BackgroundColor '#FFEB9C' -ConditionalTextColor '#9C5700'),
                        (New-ConditionalText -Text 'Tenant-wide (unrestricted)' -BackgroundColor '#F8CBAD' -ConditionalTextColor '#9C0006'))
                }
                $sheets[$k] | Export-Excel @p
            }
            $files += $x
        }
    }

    $apps = $inv.Apps
    Write-Host ('Tenant {0}: {1} apps, {2} permission rows, {3} credentials ({4} expired, {5} expiring <= {6}d), {7} access policies, {8} RBAC assignments' -f `
        $ctx.TenantId, $apps.Count, $inv.Permissions.Count, $inv.Credentials.Count, $sheets.Expired.Count, $sheets.Expiring.Count, $ExpiringDays, $aap.Count, $rbac.Count)
    foreach ($t in $script:ReportTabs) { if ($sheets[$t].Count) { Write-Host ('  {0,-14} {1}' -f $t, $sheets[$t].Count) } }
    foreach ($f in $files) { Write-Host "  -> $f" }

    [pscustomobject]@{
        TenantId       = $ctx.TenantId
        Apps           = $inv.Apps
        Permissions    = $inv.Permissions
        Credentials    = $inv.Credentials
        AccessPolicies = $aap
        ExchangeRbac   = $rbac
        EwsConfig      = $ews
        ScopeByTarget  = $scope
        Files          = $files
    }
}

#endregion

Export-ModuleMember -Function Set-AppReportConfig, Get-AppReportConfig, Connect-AppReport, Get-AppReportInventory, `
    Get-AppReportAccessPolicy, Get-AppReportExchangeRbac, Get-AppReportEwsConfig, Get-AppReportScope, Invoke-AppReport
