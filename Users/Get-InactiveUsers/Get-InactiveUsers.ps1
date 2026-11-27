#Requires -Version 5.1

<#
.SYNOPSIS
    Reports member accounts that have not signed in for more than X days.

.DESCRIPTION
    Report only: nothing is disabled or deleted.

    Uses signInActivity (last interactive, non-interactive and successful sign-in) and
    takes the most recent of the three. Accounts that never signed in count as inactive
    when they were created before the cutoff.

    The report shows per user whether the account is enabled and which licences it has,
    so you can see straight away which licences you are paying for without use.
    Guests are left out; see Get-StaleGuestUsers.ps1.

.PARAMETER TenantId
    Tenant ID or tenant domain.

.PARAMETER Days
    Days without sign-in before an account is inactive (default 90).

.PARAMETER IncludeDisabled
    Also list accounts that are already disabled.

.PARAMETER ClientId
    App registration (client) ID for app-only sign-in with a certificate. Without
    -ClientId the script signs in interactively (delegated).

.PARAMETER CertificateThumbprint
    Thumbprint of the certificate for -ClientId.

.PARAMETER OutputDirectory
    Folder for the CSV and the log. Default: current folder.

.PARAMETER Delimiter
    CSV delimiter. Default ';' (Dutch Excel).

.EXAMPLE
    .\Get-InactiveUsers.ps1 -TenantId contoso.onmicrosoft.com

.EXAMPLE
    .\Get-InactiveUsers.ps1 -TenantId $tid -Days 60 -IncludeDisabled -OutputDirectory D:\Reports

.NOTES
    Graph permissions
      Application : User.Read.All, AuditLog.Read.All, Organization.Read.All
      Delegated   : User.Read.All, AuditLog.Read.All, Organization.Read.All (Reports Reader or higher)

    signInActivity needs Entra ID P1 or P2 (included in Microsoft 365 Business Premium).
    Dependency: Microsoft.Graph.Authentication.
    Exit codes: 0 = no inactive users, 1 = inactive users found, 2 = fatal error.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [ValidateRange(1, 3650)][int]$Days = 90,
    [switch]$IncludeDisabled,
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

#region Script ------------------------------------------------------------------------------

function Get-LastSignIn($Activity) {
    $dates = foreach ($name in 'lastSignInDateTime', 'lastNonInteractiveSignInDateTime', 'lastSuccessfulSignInDateTime') {
        ConvertTo-LocalDate (Get-Value $Activity $name)
    }
    return ($dates | Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1)
}

$exitCode = 0
try {
    Connect-Tenant -Scopes 'User.Read.All', 'AuditLog.Read.All', 'Organization.Read.All'
    $script:TenantLabel = Get-TenantLabel
    Initialize-Output -ScriptName 'InactiveUsers'
    $cutoff = (Get-Date).AddDays(-$Days)
    Write-Log ("=== Inactive users in {0}: no sign-in since {1:yyyy-MM-dd} ({2} days) ===" -f $script:TenantLabel, $cutoff, $Days)

    $skus = @{}
    foreach ($s in @(Get-GraphAll -Uri 'v1.0/subscribedSkus')) { $skus[(Get-Value $s 'skuId')] = Get-Value $s 'skuPartNumber' }

    $select = 'id,displayName,userPrincipalName,userType,accountEnabled,createdDateTime,department,jobTitle,assignedLicenses,signInActivity'
    $users = @(Get-GraphAll -Uri "v1.0/users?`$select=$select")
    $members = @($users | Where-Object { (Get-Value $_ 'userType') -ne 'Guest' })
    Write-Log "$($members.Count) member account(s)"

    $rows = foreach ($u in $members) {
        $enabled = [bool](Get-Value $u 'accountEnabled')
        if (-not $enabled -and -not $IncludeDisabled) { continue }
        $last = Get-LastSignIn (Get-Value $u 'signInActivity')
        $created = ConvertTo-LocalDate (Get-Value $u 'createdDateTime')
        $inactive = if ($last) { $last -lt $cutoff } else { $created -and $created -lt $cutoff }
        if (-not $inactive) { continue }

        $licences = @(@(Get-Value $u 'assignedLicenses') | Where-Object { $_ } | ForEach-Object {
            $id = Get-Value $_ 'skuId'; if ($skus.ContainsKey($id)) { $skus[$id] } else { $id }
        })
        [pscustomobject]@{
            DisplayName       = Get-Value $u 'displayName'
            UserPrincipalName = Get-Value $u 'userPrincipalName'
            Enabled           = $enabled
            LastSignIn        = $last
            DaysSinceSignIn   = if ($last) { [int][math]::Floor(((Get-Date) - $last).TotalDays) } else { $null }
            Created           = $created
            Department        = Get-Value $u 'department'
            JobTitle          = Get-Value $u 'jobTitle'
            Licences          = $licences -join ', '
            LicenceCount      = $licences.Count
        }
    }
    $rows = @($rows | Sort-Object LastSignIn)

    $null = Export-Report -Data $rows -Name "InactiveUsers_${Days}d"
    $licensed = @($rows | Where-Object { $_.LicenceCount -gt 0 })
    Write-Log ("Inactive: {0}, of which {1} with a licence and {2} that never signed in" -f `
        $rows.Count, $licensed.Count, @($rows | Where-Object { -not $_.LastSignIn }).Count)
    $licensed | Group-Object { $_.Licences } | Sort-Object Count -Descending | ForEach-Object {
        Write-Log ('  {0,3} x {1}' -f $_.Count, $_.Name)
    }
    if ($rows.Count -gt 0) { $exitCode = 1 }
}
catch {
    if (-not $script:LogFile) { try { Initialize-Output -ScriptName 'InactiveUsers' } catch { } }
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    $exitCode = 2
}
finally {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
}
exit $exitCode

#endregion
