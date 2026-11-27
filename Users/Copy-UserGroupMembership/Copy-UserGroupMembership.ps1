#Requires -Version 5.1

<#
.SYNOPSIS
    Gives a user the same group memberships as a colleague.

.DESCRIPTION
    Reads the groups -SourceUser is a direct member of and adds -TargetUser to the ones
    the target is not in yet. Handy when someone changes jobs or joins a team.

    Skipped, with the reason in the log:
      - dynamic groups (membership comes from a rule)
      - mail-enabled security groups and distribution lists (managed in Exchange)
      - groups in -ExcludeGroups

    With -RemoveOther the target is also removed from groups the source is not in
    (same skip rules). Supports -WhatIf; use it first.

.PARAMETER TenantId
    Tenant ID or tenant domain.

.PARAMETER SourceUser
    UPN or object ID of the colleague to copy from.

.PARAMETER TargetUser
    UPN or object ID of the user who gets the memberships.

.PARAMETER ExcludeGroups
    Group names (wildcards allowed) that are never copied or removed.

.PARAMETER RemoveOther
    Also remove the target from groups the source is not a member of.

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
    .\Copy-UserGroupMembership.ps1 -TenantId contoso.onmicrosoft.com -SourceUser k.bakker@contoso.com -TargetUser j.jansen@contoso.com -WhatIf

.EXAMPLE
    .\Copy-UserGroupMembership.ps1 -TenantId $tid -SourceUser $old -TargetUser $new -RemoveOther -ExcludeGroups 'SG-Admins*'

.NOTES
    Graph permissions
      Application : User.Read.All, GroupMember.ReadWrite.All
      Delegated   : User.Read.All, GroupMember.ReadWrite.All

    Dependency: Microsoft.Graph.Authentication.
    Exit codes: 0 = done, 1 = one or more changes failed, 2 = fatal error.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [Parameter(Mandatory)][string]$SourceUser,
    [Parameter(Mandatory)][string]$TargetUser,
    [string[]]$ExcludeGroups = @(),
    [switch]$RemoveOther,
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

function Get-DirectGroups($UserId) {
    $all = Get-GraphAll -Uri "v1.0/users/$UserId/memberOf?`$select=id,displayName,groupTypes,mailEnabled,securityEnabled"
    return @($all | Where-Object { (Get-Value $_ '@odata.type') -eq '#microsoft.graph.group' })
}

function Test-Excluded($Group) {
    $name = Get-Value $Group 'displayName'
    foreach ($pattern in $ExcludeGroups) { if ($name -like $pattern) { return $true } }
    return $false
}

$exitCode = 0
$failed = 0
$changes = New-Object System.Collections.Generic.List[object]
try {
    Connect-Tenant -Scopes 'User.Read.All', 'GroupMember.ReadWrite.All'
    $script:TenantLabel = Get-TenantLabel
    Initialize-Output -ScriptName 'CopyUserGroupMembership'

    $source = Get-GraphUser -Identity $SourceUser
    $target = Get-GraphUser -Identity $TargetUser
    $targetId = Get-Value $target 'id'
    Write-Log ("=== Groups from {0} to {1} ===" -f (Get-Value $source 'userPrincipalName'), (Get-Value $target 'userPrincipalName'))

    $sourceGroups = Get-DirectGroups (Get-Value $source 'id')
    $targetGroups = Get-DirectGroups $targetId
    $targetIds = @($targetGroups | ForEach-Object { Get-Value $_ 'id' })
    $sourceIds = @($sourceGroups | ForEach-Object { Get-Value $_ 'id' })
    Write-Log "Source: $($sourceGroups.Count) group(s), target: $($targetGroups.Count) group(s)"

    $toAdd = @($sourceGroups | Where-Object { $targetIds -notcontains (Get-Value $_ 'id') })
    $toRemove = if ($RemoveOther) { @($targetGroups | Where-Object { $sourceIds -notcontains (Get-Value $_ 'id') }) } else { @() }

    foreach ($item in @($toAdd | ForEach-Object { @{ Group = $_; Action = 'Add' } }) + @($toRemove | ForEach-Object { @{ Group = $_; Action = 'Remove' } })) {
        $grp = $item.Group
        $name = Get-Value $grp 'displayName'
        $gid = Get-Value $grp 'id'
        $reason = Test-GroupWritable $grp
        if (-not $reason -and (Test-Excluded $grp)) { $reason = 'excluded' }

        $result = if ($reason) { "Skipped: $reason" } else { 'WhatIf' }
        if (-not $reason -and $PSCmdlet.ShouldProcess($name, "$($item.Action) $(Get-Value $target 'userPrincipalName')")) {
            try {
                if ($item.Action -eq 'Add') {
                    Invoke-GraphApi -Uri "v1.0/groups/$gid/members/`$ref" -Method POST `
                        -Body @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$targetId" } | Out-Null
                }
                else {
                    Invoke-GraphApi -Uri "v1.0/groups/$gid/members/$targetId/`$ref" -Method DELETE | Out-Null
                }
                $result = 'Done'
            }
            catch {
                $result = "Failed: $($_.Exception.Message)"
                $failed++
            }
        }
        Write-Log ("  {0,-6} {1}: {2}" -f $item.Action, $name, $result) $(if ($result -like 'Failed*') { 'WARN' } else { 'INFO' })
        $changes.Add([pscustomobject]@{ Action = $item.Action; Group = $name; GroupId = $gid; Result = $result })
    }

    if ($changes.Count -eq 0) { Write-Log 'Nothing to change: the target already has the same groups.' }
    $null = Export-Report -Data $changes -Name 'CopyUserGroupMembership'
    if ($failed -gt 0) { $exitCode = 1 }
}
catch {
    if (-not $script:LogFile) { try { Initialize-Output -ScriptName 'CopyUserGroupMembership' } catch { } }
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    $exitCode = 2
}
finally {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
}
exit $exitCode

#endregion
