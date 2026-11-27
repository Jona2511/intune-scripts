#Requires -Version 5.1

<#
.SYNOPSIS
    Offboards a Microsoft 365 user: blocks sign-in, ends sessions, removes groups and licences.

.DESCRIPTION
    Steps, in this order:

      1. Blocks sign-in (accountEnabled = false).
      2. Revokes all refresh tokens and sessions, so open apps have to sign in again.
      3. Resets the password to a random one (-ResetPassword).
      4. Disables the user's registered and joined devices in Entra ID (-DisableDevices).
      5. Removes the user from all groups that Graph can change (-RemoveGroups). Dynamic
         and mail-enabled groups are left alone and listed in the log.
      6. Removes directly assigned licences (-RemoveLicenses). Licences that come from a
         group disappear with step 5.

    Everything the script removes is written to a CSV first, so you can put it back.
    Supports -WhatIf: run that first and read the log.

    Convert the mailbox to a shared mailbox in Exchange BEFORE you use -RemoveLicenses.
    Without a licence, Exchange deletes the mailbox after 30 days.

.PARAMETER TenantId
    Tenant ID or tenant domain.

.PARAMETER User
    UPN or object ID of the user.

.PARAMETER ResetPassword
    Set a random password. Not shown or logged.

.PARAMETER DisableDevices
    Disable the user's devices in Entra ID.

.PARAMETER RemoveGroups
    Remove the user from groups.

.PARAMETER RemoveLicenses
    Remove directly assigned licences.

.PARAMETER ClientId
    App registration (client) ID for app-only sign-in with a certificate. Without
    -ClientId the script signs in interactively (delegated).

.PARAMETER CertificateThumbprint
    Thumbprint of the certificate for -ClientId.

.PARAMETER OutputDirectory
    Folder for the log and the CSV. Default: current folder.

.PARAMETER Delimiter
    CSV delimiter. Default ';' (Dutch Excel).

.EXAMPLE
    .\Invoke-UserOffboarding.ps1 -TenantId contoso.onmicrosoft.com -User j.jansen@contoso.com -RemoveGroups -RemoveLicenses -WhatIf

.EXAMPLE
    .\Invoke-UserOffboarding.ps1 -TenantId $tid -User j.jansen@contoso.com -ResetPassword -DisableDevices -RemoveGroups -RemoveLicenses

.NOTES
    Graph permissions
      Application : User.ReadWrite.All, GroupMember.ReadWrite.All, Device.ReadWrite.All (for -DisableDevices),
                    User-PasswordProfile.ReadWrite.All (for -ResetPassword)
      Delegated   : User.ReadWrite.All, GroupMember.ReadWrite.All, Directory.AccessAsUser.All
                    (signed in as User Administrator; Cloud Device Administrator for -DisableDevices)

    Dependency: Microsoft.Graph.Authentication.
    Exit codes: 0 = done, 1 = one or more steps failed, 2 = fatal error.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$User,
    [switch]$ResetPassword,
    [switch]$DisableDevices,
    [switch]$RemoveGroups,
    [switch]$RemoveLicenses,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$OutputDirectory = (Get-Location).Path,
    [string]$Delimiter = ';'
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
$failed = 0
$record = New-Object System.Collections.Generic.List[object]

function Invoke-Step {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][string]$Target, [Parameter(Mandatory)][string]$Action, [Parameter(Mandatory)][scriptblock]$Do)
    if (-not $PSCmdlet.ShouldProcess($Target, $Action)) { Write-Log "  WhatIf: $Action ($Target)"; return }
    try {
        & $Do
        Write-Log "  $Action ($Target)"
    }
    catch {
        Write-Log "  FAILED: $Action ($Target): $($_.Exception.Message)" 'WARN'
        $script:failed++
    }
}

try {
    $scopes = @('User.ReadWrite.All', 'GroupMember.ReadWrite.All')
    if ($DisableDevices) { $scopes += 'Device.ReadWrite.All' }
    if ($ResetPassword) { $scopes += 'Directory.AccessAsUser.All' }
    Connect-Tenant -Scopes $scopes
    $script:TenantLabel = Get-TenantLabel
    Initialize-Output -ScriptName 'UserOffboarding'

    $u = Get-GraphUser -Identity $User -Select 'id,displayName,userPrincipalName,accountEnabled,assignedLicenses,licenseAssignmentStates'
    $uid = Get-Value $u 'id'
    $upn = Get-Value $u 'userPrincipalName'
    Write-Log "=== Offboarding $upn ($(Get-Value $u 'displayName')) ==="

    $groups = @(Get-GraphAll -Uri "v1.0/users/$uid/memberOf?`$select=id,displayName,groupTypes,mailEnabled,securityEnabled" |
        Where-Object { (Get-Value $_ '@odata.type') -eq '#microsoft.graph.group' })
    $skus = @{}
    foreach ($s in @(Get-GraphAll -Uri 'v1.0/subscribedSkus')) { $skus[(Get-Value $s 'skuId')] = Get-Value $s 'skuPartNumber' }
    $directLicenses = @(@(Get-Value $u 'licenseAssignmentStates') |
        Where-Object { $_ -and -not (Get-Value $_ 'assignedByGroup') } |
        ForEach-Object { Get-Value $_ 'skuId' } | Sort-Object -Unique)
    $devices = if ($DisableDevices) { @(Get-GraphAll -Uri "v1.0/users/$uid/ownedDevices?`$select=id,displayName,accountEnabled") } else { @() }

    foreach ($g in $groups) {
        $record.Add([pscustomobject]@{ Type = 'Group'; Name = Get-Value $g 'displayName'; Id = Get-Value $g 'id'; Note = (Test-GroupWritable $g) })
    }
    foreach ($l in $directLicenses) {
        $record.Add([pscustomobject]@{ Type = 'Licence'; Name = $skus[$l]; Id = $l; Note = 'direct' })
    }
    foreach ($d in $devices) {
        $record.Add([pscustomobject]@{ Type = 'Device'; Name = Get-Value $d 'displayName'; Id = Get-Value $d 'id'; Note = "enabled=$(Get-Value $d 'accountEnabled')" })
    }
    $null = Export-Report -Data $record -Name "Offboarding_before_$($upn.Split('@')[0])"

    Write-Log '1. Block sign-in'
    Invoke-Step $upn 'Block sign-in' { Invoke-GraphApi -Uri "v1.0/users/$uid" -Method PATCH -Body @{ accountEnabled = $false } | Out-Null }

    Write-Log '2. Revoke sessions'
    Invoke-Step $upn 'Revoke sign-in sessions' { Invoke-GraphApi -Uri "v1.0/users/$uid/revokeSignInSessions" -Method POST | Out-Null }

    if ($ResetPassword) {
        Write-Log '3. Reset password'
        Invoke-Step $upn 'Reset password' {
            $body = @{ passwordProfile = @{ password = (New-RandomPassword -Length 24); forceChangePasswordNextSignIn = $true } }
            Invoke-GraphApi -Uri "v1.0/users/$uid" -Method PATCH -Body $body | Out-Null
        }
    }

    if ($DisableDevices) {
        Write-Log "4. Disable devices ($($devices.Count))"
        foreach ($d in $devices) {
            if (-not (Get-Value $d 'accountEnabled')) { continue }
            $did = Get-Value $d 'id'
            Invoke-Step (Get-Value $d 'displayName') 'Disable device' { Invoke-GraphApi -Uri "v1.0/devices/$did" -Method PATCH -Body @{ accountEnabled = $false } | Out-Null }
        }
    }

    if ($RemoveGroups) {
        Write-Log "5. Remove from groups ($($groups.Count))"
        foreach ($g in $groups) {
            $name = Get-Value $g 'displayName'
            $reason = Test-GroupWritable $g
            if ($reason) { Write-Log "  Left alone: $name ($reason)" 'WARN'; continue }
            $gid = Get-Value $g 'id'
            Invoke-Step $name 'Remove from group' { Invoke-GraphApi -Uri "v1.0/groups/$gid/members/$uid/`$ref" -Method DELETE | Out-Null }
        }
    }

    if ($RemoveLicenses) {
        Write-Log "6. Remove direct licences ($($directLicenses.Count))"
        if ($directLicenses.Count -gt 0) {
            $names = ($directLicenses | ForEach-Object { $skus[$_] }) -join ', '
            Invoke-Step $upn "Remove licences: $names" {
                Invoke-GraphApi -Uri "v1.0/users/$uid/assignLicense" -Method POST -Body @{ addLicenses = @(); removeLicenses = @($directLicenses) } | Out-Null
            }
        }
    }

    Write-Log "Done, $failed step(s) failed. Not done by this script: mailbox (convert to shared, forwarding, out of office) and OneDrive access for the manager."
    if ($failed -gt 0) { $exitCode = 1 }
}
catch {
    if (-not $script:LogFile) { try { Initialize-Output -ScriptName 'UserOffboarding' } catch { } }
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    $exitCode = 2
}
finally {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
}
exit $exitCode

#endregion
