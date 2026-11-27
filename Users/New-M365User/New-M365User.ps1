#Requires -Version 5.1

<#
.SYNOPSIS
    Creates a Microsoft 365 user with licence, groups and manager in one go.

.DESCRIPTION
    Onboarding for a new employee:

      1. Creates the user with a random temporary password (must be changed at first sign-in).
      2. Sets the manager (-Manager).
      3. Assigns a licence directly (-LicenseSku). Skip this when you use group-based
         licensing and add the licence group to -Groups instead.
      4. Adds the user to groups (-Groups) and/or to the same groups as a colleague
         (-CopyGroupsFrom). Dynamic and mail-enabled groups are skipped, because Graph
         cannot change them.

    The temporary password is shown once on screen and is never written to the log.
    Supports -WhatIf.

.PARAMETER TenantId
    Tenant ID or tenant domain.

.PARAMETER UserPrincipalName
    UPN of the new user, e.g. j.jansen@contoso.com.

.PARAMETER GivenName
    First name.

.PARAMETER Surname
    Last name.

.PARAMETER DisplayName
    Display name. Default: "<GivenName> <Surname>".

.PARAMETER JobTitle
    Job title.

.PARAMETER Department
    Department.

.PARAMETER UsageLocation
    Two-letter country code, needed for licences. Default NL.

.PARAMETER Manager
    UPN or object ID of the manager.

.PARAMETER LicenseSku
    SkuPartNumber of the licence, e.g. SPB (Microsoft 365 Business Premium) or O365_BUSINESS_STANDARD.

.PARAMETER Groups
    Group names or IDs to add the user to.

.PARAMETER CopyGroupsFrom
    UPN of a colleague; the new user gets the same (assigned) group memberships.

.PARAMETER ClientId
    App registration (client) ID for app-only sign-in with a certificate. Without
    -ClientId the script signs in interactively (delegated).

.PARAMETER CertificateThumbprint
    Thumbprint of the certificate for -ClientId.

.PARAMETER OutputDirectory
    Folder for the log. Default: current folder.

.EXAMPLE
    .\New-M365User.ps1 -TenantId contoso.onmicrosoft.com -UserPrincipalName j.jansen@contoso.com `
        -GivenName Jan -Surname Jansen -JobTitle 'Sales' -Manager p.devries@contoso.com `
        -LicenseSku SPB -CopyGroupsFrom k.bakker@contoso.com -WhatIf

.NOTES
    Graph permissions
      Application : User.ReadWrite.All, GroupMember.ReadWrite.All, Organization.Read.All
      Delegated   : User.ReadWrite.All, GroupMember.ReadWrite.All, Organization.Read.All
                    (signed in as User Administrator or higher)

    Dependency: Microsoft.Graph.Authentication.
    Exit codes: 0 = done, 1 = user created but one or more steps failed, 2 = fatal error.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')][string]$UserPrincipalName,
    [Parameter(Mandatory)][string]$GivenName,
    [Parameter(Mandatory)][string]$Surname,
    [string]$DisplayName,
    [string]$JobTitle,
    [string]$Department,
    [ValidatePattern('^[A-Z]{2}$')][string]$UsageLocation = 'NL',
    [string]$Manager,
    [string]$LicenseSku,
    [string[]]$Groups = @(),
    [string]$CopyGroupsFrom,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$OutputDirectory = (Get-Location).Path
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Common helpers (identical in all tenant scripts) --------------------------------------

$script:LogFile     = $null
$script:LogBuffer   = New-Object System.Collections.Generic.List[string]
$script:TenantLabel = $TenantId

function Write-Log {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    switch ($Level) {
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }
    if ($script:LogFile) { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -WhatIf:$false }
    else { $script:LogBuffer.Add($line) }
}

function Initialize-Output {
    param([Parameter(Mandatory)][string]$ScriptName)
    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        New-Item -Path $OutputDirectory -ItemType Directory -Force -WhatIf:$false | Out-Null
    }
    $script:LogFile = Join-Path $OutputDirectory ('{0}_{1}_{2:yyyyMMdd_HHmmss}.log' -f $ScriptName, $script:TenantLabel, (Get-Date))
    if ($script:LogBuffer.Count -gt 0) {
        Add-Content -LiteralPath $script:LogFile -Value $script:LogBuffer -Encoding UTF8 -WhatIf:$false
        $script:LogBuffer.Clear()
    }
}

function Get-Value {
    <# Safe property access under StrictMode: $null when the property does not exist. #>
    param($Object, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $Object) { return $null }
    $p = $Object.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $null
}

function ConvertTo-LocalDate {
    <# Graph dates arrive as string or DateTime; '0001-01-01' means "never". #>
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    $d = if ($Value -is [datetime]) { $Value }
         else { [datetime]::Parse("$Value", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind) }
    if ($d.Year -le 1601) { return $null }
    return $d.ToLocalTime()
}

function Connect-Tenant {
    param([Parameter(Mandatory)][string[]]$Scopes)
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw 'Module Microsoft.Graph.Authentication is missing: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    # ContextScope Process: no token reuse between runs against different tenants
    if ($ClientId) {
        if (-not $CertificateThumbprint) { throw '-CertificateThumbprint is required together with -ClientId.' }
        Write-Log "Connecting to $TenantId (app-only, client $ClientId)"
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint `
            -ContextScope Process -NoWelcome
    }
    else {
        Write-Log "Connecting to $TenantId (delegated, scopes: $($Scopes -join ', '))"
        Connect-MgGraph -TenantId $TenantId -Scopes $Scopes -ContextScope Process -NoWelcome
    }
}

function Invoke-GraphApi {
    <# Invoke-MgGraphRequest with retry on 429/503/504, honoring Retry-After. #>
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Method = 'GET',
        $Body,
        [int]$MaxRetries = 6
    )
    for ($attempt = 1; ; $attempt++) {
        try {
            $req = @{ Uri = $Uri; Method = $Method; OutputType = 'PSObject'; ErrorAction = 'Stop' }
            if ($null -ne $Body) {
                $req.Body        = ($Body | ConvertTo-Json -Depth 10 -Compress)
                $req.ContentType = 'application/json'
            }
            return Invoke-MgGraphRequest @req
        }
        catch {
            $status = $null
            $wait   = $null
            $respProp = $_.Exception.PSObject.Properties['Response']
            if ($respProp -and $respProp.Value) {
                try { $status = [int]$respProp.Value.StatusCode } catch { }
                try {
                    $ra = $respProp.Value.Headers.RetryAfter
                    if ($ra -and $ra.Delta) { $wait = [int][math]::Ceiling($ra.Delta.TotalSeconds) }
                } catch { }
            }
            if (-not $status -and $_.Exception.Message -match '\b(429|503|504)\b') { $status = [int]$Matches[1] }

            if ($status -in 429, 503, 504 -and $attempt -le $MaxRetries) {
                if (-not $wait) { $wait = [int][math]::Min(60, [math]::Pow(2, $attempt)) }
                Write-Log "  Graph returned $status; retry $attempt/$MaxRetries in $wait s" 'WARN'
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }
}

function Get-GraphAll {
    <# GET with @odata.nextLink paging; returns all items of 'value'. #>
    param([Parameter(Mandatory)][string]$Uri)
    $items = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    while ($next) {
        $resp = Invoke-GraphApi -Uri $next
        $val = Get-Value $resp 'value'
        if ($null -ne $val) { foreach ($v in @($val)) { $items.Add($v) } }
        $next = Get-Value $resp '@odata.nextLink'
    }
    return , $items.ToArray()
}

function Get-TenantLabel {
    try {
        $org = Get-GraphAll -Uri 'v1.0/organization?$select=verifiedDomains'
        $domain = @(Get-Value @($org)[0] 'verifiedDomains') |
            Where-Object { Get-Value $_ 'isDefault' } | Select-Object -First 1
        if ($domain) { return ((Get-Value $domain 'name') -replace '[\\/:*?"<>|]', '_') }
    }
    catch { Write-Log 'Could not read the tenant domain (Organization.Read.All missing?); using the tenant ID in file names.' 'WARN' }
    return ($TenantId -replace '[\\/:*?"<>|]', '_')
}

function Export-Report {
    param([AllowNull()][object[]]$Data, [Parameter(Mandatory)][string]$Name)
    $path = Join-Path $OutputDirectory ('{0}_{1}_{2:yyyyMMdd}.csv' -f $Name, $script:TenantLabel, (Get-Date))
    $rows = @($Data | Where-Object { $null -ne $_ })
    if ($rows.Count -gt 0) {
        $rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8 -Delimiter $Delimiter -WhatIf:$false
    }
    else {
        Set-Content -LiteralPath $path -Value '' -Encoding UTF8 -WhatIf:$false
    }
    Write-Log "Exported $($rows.Count) row(s) to $path"
    return $path
}

#endregion

#region User helpers -------------------------------------------------------------------------

function Get-GraphUser {
    param([Parameter(Mandatory)][string]$Identity, [string]$Select = 'id,displayName,userPrincipalName,accountEnabled')
    $id = [uri]::EscapeDataString($Identity)
    return Invoke-GraphApi -Uri "v1.0/users/${id}?`$select=$Select"
}

function Get-GraphGroup {
    param([Parameter(Mandatory)][string]$NameOrId)
    $select = 'id,displayName,groupTypes,mailEnabled,securityEnabled'
    if ($NameOrId -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
        return Invoke-GraphApi -Uri "v1.0/groups/${NameOrId}?`$select=$select"
    }
    $escaped = $NameOrId -replace "'", "''"
    $filter = [uri]::EscapeDataString("displayName eq '$escaped'")
    $found = @(Get-GraphAll -Uri "v1.0/groups?`$filter=$filter&`$select=$select")
    if ($found.Count -eq 0) { throw "Group '$NameOrId' not found." }
    if ($found.Count -gt 1) { throw "Group name '$NameOrId' is not unique ($($found.Count) groups). Use the group ID." }
    return $found[0]
}

function Test-GroupWritable {
    <# Graph cannot change members of dynamic groups or mail-enabled groups (those are managed in Exchange). #>
    param([Parameter(Mandatory)]$Group)
    if (@(Get-Value $Group 'groupTypes') -contains 'DynamicMembership') { return 'dynamic membership' }
    if (Get-Value $Group 'mailEnabled') {
        if (@(Get-Value $Group 'groupTypes') -contains 'Unified') { return $null }
        return 'mail-enabled, manage it in Exchange'
    }
    return $null
}

function New-RandomPassword {
    param([ValidateRange(12, 64)][int]$Length = 16)
    $sets = 'ABCDEFGHJKLMNPQRSTUVWXYZ', 'abcdefghijkmnopqrstuvwxyz', '23456789', '!@#$%*+-=?'
    $all = -join $sets
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $bytes = New-Object byte[] ($Length * 2)
    $rng.GetBytes($bytes)
    $chars = New-Object System.Collections.Generic.List[char]
    for ($i = 0; $i -lt $Length; $i++) {
        $set = if ($i -lt $sets.Count) { $sets[$i] } else { $all }
        $chars.Add($set[$bytes[$i] % $set.Length])
    }
    for ($i = $chars.Count - 1; $i -gt 0; $i--) {
        $j = $bytes[$Length + $i] % ($i + 1)
        $tmp = $chars[$i]; $chars[$i] = $chars[$j]; $chars[$j] = $tmp
    }
    $rng.Dispose()
    return -join $chars
}

#endregion

#region Script ------------------------------------------------------------------------------

$exitCode = 0
$warnings = 0
try {
    Connect-Tenant -Scopes 'User.ReadWrite.All', 'GroupMember.ReadWrite.All', 'Organization.Read.All'
    $script:TenantLabel = Get-TenantLabel
    Initialize-Output -ScriptName 'NewM365User'
    if (-not $DisplayName) { $DisplayName = "$GivenName $Surname" }
    Write-Log "=== New user $UserPrincipalName ($DisplayName) in $($script:TenantLabel) ==="

    $existing = $null
    try { $existing = Get-GraphUser -Identity $UserPrincipalName } catch { }
    if ($existing) { throw "User $UserPrincipalName already exists." }

    # Look everything up first, so a typo stops the script before the user exists
    $managerObj = if ($Manager) { Get-GraphUser -Identity $Manager } else { $null }

    $sku = $null
    if ($LicenseSku) {
        $sku = @(Get-GraphAll -Uri 'v1.0/subscribedSkus') | Where-Object { (Get-Value $_ 'skuPartNumber') -eq $LicenseSku } | Select-Object -First 1
        if (-not $sku) { throw "Licence $LicenseSku is not in this tenant. Check Get-LicenseOverview.ps1 for the SkuPartNumbers." }
        $free = [int](Get-Value (Get-Value $sku 'prepaidUnits') 'enabled') - [int](Get-Value $sku 'consumedUnits')
        if ($free -lt 1) { throw "No free $LicenseSku licences left (0 of $(Get-Value (Get-Value $sku 'prepaidUnits') 'enabled'))." }
        Write-Log "Licence ${LicenseSku}: $free free"
    }

    $targetGroups = @{}
    foreach ($g in $Groups) {
        $grp = Get-GraphGroup -NameOrId $g
        $targetGroups[(Get-Value $grp 'id')] = $grp
    }
    if ($CopyGroupsFrom) {
        $source = Get-GraphUser -Identity $CopyGroupsFrom
        $memberOf = Get-GraphAll -Uri "v1.0/users/$(Get-Value $source 'id')/memberOf?`$select=id,displayName,groupTypes,mailEnabled,securityEnabled"
        $copied = @($memberOf | Where-Object { (Get-Value $_ '@odata.type') -eq '#microsoft.graph.group' })
        foreach ($grp in $copied) { $targetGroups[(Get-Value $grp 'id')] = $grp }
        Write-Log "Copying from ${CopyGroupsFrom}: $($copied.Count) group(s)"
    }

    $password = New-RandomPassword
    $body = @{
        accountEnabled    = $true
        displayName       = $DisplayName
        givenName         = $GivenName
        surname           = $Surname
        userPrincipalName = $UserPrincipalName
        mailNickname      = ($UserPrincipalName.Split('@')[0] -replace '[^A-Za-z0-9._-]', '')
        usageLocation     = $UsageLocation
        passwordProfile   = @{ password = $password; forceChangePasswordNextSignIn = $true }
    }
    if ($JobTitle)   { $body.jobTitle = $JobTitle }
    if ($Department) { $body.department = $Department }

    if (-not $PSCmdlet.ShouldProcess($UserPrincipalName, 'Create user')) {
        Write-Log "WhatIf: would create $UserPrincipalName, licence '$LicenseSku', manager '$Manager', $($targetGroups.Count) group(s)"
        foreach ($grp in $targetGroups.Values) {
            $reason = Test-GroupWritable $grp
            Write-Log ("  {0}{1}" -f (Get-Value $grp 'displayName'), $(if ($reason) { " (skipped: $reason)" } else { '' }))
        }
        return
    }

    $user = Invoke-GraphApi -Uri 'v1.0/users' -Method POST -Body $body
    $userId = Get-Value $user 'id'
    Write-Log "Created $UserPrincipalName ($userId)"

    # A new user is not visible everywhere straight away
    for ($i = 1; $i -le 10; $i++) {
        try { $null = Get-GraphUser -Identity $userId; break } catch { Start-Sleep -Seconds 3 }
    }

    if ($managerObj) {
        try {
            Invoke-GraphApi -Uri "v1.0/users/$userId/manager/`$ref" -Method PUT `
                -Body @{ '@odata.id' = "https://graph.microsoft.com/v1.0/users/$(Get-Value $managerObj 'id')" } | Out-Null
            Write-Log "Manager: $(Get-Value $managerObj 'userPrincipalName')"
        }
        catch { Write-Log "Manager not set: $($_.Exception.Message)" 'WARN'; $warnings++ }
    }

    if ($sku) {
        try {
            Invoke-GraphApi -Uri "v1.0/users/$userId/assignLicense" -Method POST `
                -Body @{ addLicenses = @(@{ skuId = (Get-Value $sku 'skuId') }); removeLicenses = @() } | Out-Null
            Write-Log "Licence assigned: $LicenseSku"
        }
        catch { Write-Log "Licence not assigned: $($_.Exception.Message)" 'WARN'; $warnings++ }
    }

    foreach ($grp in $targetGroups.Values) {
        $name = Get-Value $grp 'displayName'
        $reason = Test-GroupWritable $grp
        if ($reason) { Write-Log "  Skipped $name ($reason)" 'WARN'; continue }
        try {
            Invoke-GraphApi -Uri "v1.0/groups/$(Get-Value $grp 'id')/members/`$ref" -Method POST `
                -Body @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$userId" } | Out-Null
            Write-Log "  Added to $name"
        }
        catch { Write-Log "  Not added to ${name}: $($_.Exception.Message)" 'WARN'; $warnings++ }
    }

    Write-Host ''
    Write-Host "Temporary password for ${UserPrincipalName}: $password" -ForegroundColor Cyan
    Write-Host 'The user has to change it at first sign-in. It is not in the log.'
    Write-Log "Done with $warnings warning(s)"
    if ($warnings -gt 0) { $exitCode = 1 }
}
catch {
    if (-not $script:LogFile) { try { Initialize-Output -ScriptName 'NewM365User' } catch { } }
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    $exitCode = 2
}
finally {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
}
exit $exitCode

#endregion
