#Requires -Version 5.1

<#
.SYNOPSIS
    Checks whether one or more IP addresses (e.g. the office IP) fall inside a trusted
    Conditional Access named location, and optionally fixes it.

.DESCRIPTION
    - Reads all IP named locations (/identity/conditionalAccess/namedLocations, v1.0).
    - For every -IpAddress (IPv4/IPv6, single address or CIDR) it determines which named
      locations contain it (real CIDR math, not string comparison) and whether any of them
      is marked as trusted. A CIDR input only counts as contained when the whole range lies
      inside the location's range.
    - Lists the Conditional Access policies that include or exclude those locations, plus
      policies that use 'AllTrusted' when the IP is in a trusted location.

    With -Fix (supports -WhatIf):
      - With -LocationName: IPs that are not in a trusted location are added to that named
        location (as /32 or /128, or the network of the given CIDR), and the location is
        marked trusted. If the location does not exist it is created as trusted.
      - Without -LocationName: untrusted named locations that already contain the IP are
        marked trusted. Note: that trusts the complete range of that location.
      Afterwards the IPs are evaluated again.

    Not covered: the legacy "trusted IPs" in the per-user MFA service settings.

    Output: NamedLocations_<domain>_<date>.csv

.PARAMETER TenantId
    Tenant ID or initial domain.

.PARAMETER IpAddress
    One or more IP addresses or CIDR ranges, e.g. '203.0.113.10', '203.0.113.0/28', '2001:db8::/48'.

.PARAMETER ClientId
    App registration (client) ID for app-only authentication. Without it: interactive (delegated).

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only authentication.

.PARAMETER Fix
    Repair (see description). Use with -WhatIf first.

.PARAMETER LocationName
    Named location to add the IPs to when using -Fix.

.PARAMETER SelfTest
    Runs the built-in CIDR tests locally, without connecting to a tenant.

.EXAMPLE
    .\Test-NamedLocationTrusted.ps1 -TenantId contoso.onmicrosoft.com -IpAddress 203.0.113.10

.EXAMPLE
    .\Test-NamedLocationTrusted.ps1 -TenantId $tid -IpAddress 203.0.113.10,2001:db8:5::/48 -Fix -LocationName 'Kantoor' -WhatIf

.EXAMPLE
    .\Test-NamedLocationTrusted.ps1 -SelfTest

.NOTES
    Required Graph permissions (application or delegated):
      Policy.Read.All                       named locations and CA policies (check)
      Policy.ReadWrite.ConditionalAccess    only for -Fix
      Organization.Read.All                 tenant name (optional)
    Delegated -Fix: the user also needs Conditional Access Administrator or Security Administrator.

    Dependency: Microsoft.Graph.Authentication.
    Exit codes: 0 = every IP is in a trusted named location, 1 = at least one is not,
    2 = fatal error (or a failed self-test).
#>
[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Check')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Check')][string]$TenantId,
    [Parameter(Mandatory, ParameterSetName = 'Check')][string[]]$IpAddress,
    [Parameter(ParameterSetName = 'Check')][string]$ClientId,
    [Parameter(ParameterSetName = 'Check')][string]$CertificateThumbprint,
    [Parameter(ParameterSetName = 'Check')][switch]$Fix,
    [Parameter(ParameterSetName = 'Check')][string]$LocationName,

    [Parameter(Mandatory, ParameterSetName = 'SelfTest')][switch]$SelfTest,

    [string]$OutputDirectory = (Get-Location).Path,
    [string]$Delimiter = ';'
)

# StrictMode 1 only: Graph objects may omit properties, and level 2+ throws on that.
Set-StrictMode -Version 1.0
$ErrorActionPreference = 'Stop'

#region CIDR math --------------------------------------------------------------------------------

function ConvertTo-IpPrefix {
    <# Parses 'a.b.c.d', 'a.b.c.d/n', IPv6 or IPv6/n into family, address bytes and prefix length. #>
    param([Parameter(Mandatory)][string]$Cidr)

    $text  = $Cidr.Trim()
    $parts = $text.Split('/')
    if ($parts.Count -gt 2) { throw "Invalid address: $Cidr" }
    if ($parts[0] -notmatch ':' -and $parts[0] -notmatch '^\d{1,3}(\.\d{1,3}){3}$') {
        throw "Invalid IPv4 address (use four octets): $Cidr"
    }
    $ip = $null
    if (-not [System.Net.IPAddress]::TryParse($parts[0], [ref]$ip)) { throw "Invalid IP address: $Cidr" }
    if ($ip.IsIPv4MappedToIPv6) { $ip = $ip.MapToIPv4() }

    $bytes = $ip.GetAddressBytes()
    $max   = $bytes.Length * 8
    $len   = $max
    if ($parts.Count -eq 2) {
        if ($parts[1] -notmatch '^\d{1,3}$') { throw "Invalid prefix length: $Cidr" }
        $len = [int]$parts[1]
    }
    if ($len -lt 0 -or $len -gt $max) { throw "Prefix length out of range: $Cidr" }

    [pscustomobject]@{ Family = $ip.AddressFamily; Bytes = $bytes; PrefixLength = $len; Text = $text }
}

function Test-PrefixBitsEqual([byte[]]$A, [byte[]]$B, [int]$Bits) {
    $full = [int][math]::Floor($Bits / 8)
    $rest = $Bits % 8
    for ($i = 0; $i -lt $full; $i++) { if ($A[$i] -ne $B[$i]) { return $false } }
    if ($rest -gt 0) {
        $mask = [byte]((0xFF -shl (8 - $rest)) -band 0xFF)
        if (($A[$full] -band $mask) -ne ($B[$full] -band $mask)) { return $false }
    }
    return $true
}

function Test-CidrContains {
    <# True if the whole of $Address (IP or CIDR) lies within $Network (CIDR). #>
    param([Parameter(Mandatory)][string]$Network, [Parameter(Mandatory)][string]$Address)
    $n = ConvertTo-IpPrefix $Network
    $a = ConvertTo-IpPrefix $Address
    if ($n.Family -ne $a.Family) { return $false }
    if ($a.PrefixLength -lt $n.PrefixLength) { return $false }
    return (Test-PrefixBitsEqual $n.Bytes $a.Bytes $n.PrefixLength)
}

function Get-NetworkCidr {
    <# Returns the CIDR with host bits cleared, e.g. 192.168.1.77/24 -> 192.168.1.0/24. #>
    param([Parameter(Mandatory)][string]$Cidr)
    $p = ConvertTo-IpPrefix $Cidr
    $bytes = [byte[]]$p.Bytes.Clone()
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        $bitsInByte = [math]::Max(0, [math]::Min(8, $p.PrefixLength - ($i * 8)))
        $mask = [byte]((0xFF -shl (8 - $bitsInByte)) -band 0xFF)
        $bytes[$i] = $bytes[$i] -band $mask
    }
    '{0}/{1}' -f ([System.Net.IPAddress]::new($bytes)).ToString(), $p.PrefixLength
}

function Invoke-SelfTest {
    $cases = @(
        @{ N = '10.0.0.0/8';        A = '10.1.2.3';           E = $true  },
        @{ N = '10.0.0.0/8';        A = '11.0.0.1';           E = $false },
        @{ N = '192.168.1.0/24';    A = '192.168.1.0/25';     E = $true  },
        @{ N = '192.168.1.0/25';    A = '192.168.1.0/24';     E = $false },
        @{ N = '203.0.113.17/32';   A = '203.0.113.17';       E = $true  },
        @{ N = '203.0.113.16/28';   A = '203.0.113.31';       E = $true  },
        @{ N = '203.0.113.16/28';   A = '203.0.113.32';       E = $false },
        @{ N = '172.16.0.0/12';     A = '172.31.255.255';     E = $true  },
        @{ N = '172.16.0.0/12';     A = '172.32.0.0';         E = $false },
        @{ N = '0.0.0.0/0';         A = '8.8.8.8';            E = $true  },
        @{ N = '2001:db8::/32';     A = '2001:db8:1234::1';   E = $true  },
        @{ N = '2001:db8::/32';     A = '2001:db9::1';        E = $false },
        @{ N = '2001:db8:abcd::/47'; A = '2001:db8:abcd:1::1'; E = $true },
        @{ N = '2001:db8:abcd::/47'; A = '2001:db8:abce::1';  E = $false },
        @{ N = '2001:db8::/32';     A = '10.0.0.1';           E = $false },
        @{ N = '10.0.0.0/8';        A = '::ffff:10.1.1.1';    E = $true  }
    )
    $failed = 0
    foreach ($c in $cases) {
        $r = Test-CidrContains -Network $c.N -Address $c.A
        $ok = ($r -eq $c.E)
        if (-not $ok) { $failed++ }
        Write-Host ('{0}  {1,-22} contains {2,-20} -> {3} (expected {4})' -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $c.N, $c.A, $r, $c.E)
    }

    $net = @(
        @{ I = '192.168.1.77/24'; E = '192.168.1.0/24' },
        @{ I = '203.0.113.10';    E = '203.0.113.10/32' },
        @{ I = '10.20.30.40/13';  E = '10.16.0.0/13' },
        @{ I = '2001:db8::1/64';  E = '2001:db8::/64' }
    )
    foreach ($c in $net) {
        $r = Get-NetworkCidr $c.I
        $ok = ($r -eq $c.E)
        if (-not $ok) { $failed++ }
        Write-Host ('{0}  network of {1,-18} -> {2} (expected {3})' -f $(if ($ok) { 'PASS' } else { 'FAIL' }), $c.I, $r, $c.E)
    }

    foreach ($bad in '10.1', '300.1.1.1', '10.0.0.0/33', '2001:db8::/129', 'abc') {
        $threw = $false
        try { $null = ConvertTo-IpPrefix $bad } catch { $threw = $true }
        if (-not $threw) { $failed++ }
        Write-Host ('{0}  invalid input {1,-16} rejected' -f $(if ($threw) { 'PASS' } else { 'FAIL' }), $bad)
    }

    Write-Host ''
    if ($failed -eq 0) { Write-Host 'All self-tests passed.' -ForegroundColor Green; return $true }
    Write-Host "$failed self-test(s) failed." -ForegroundColor Red
    return $false
}

#endregion

if ($SelfTest) {
    if (Invoke-SelfTest) { exit 0 } else { exit 2 }
}

$Scopes = @('Policy.Read.All', 'Organization.Read.All')
if ($Fix) { $Scopes += 'Policy.ReadWrite.ConditionalAccess' }

#region Shared Graph helpers (identical in all tenant scripts) --------------------------------

$script:LogFile = $null
$script:Tenant  = $null

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
    if ($script:LogFile) {
        Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -WhatIf:$false
    }
}

function Initialize-Output([string]$Prefix) {
    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        New-Item -Path $OutputDirectory -ItemType Directory -Force -WhatIf:$false | Out-Null
    }
    $script:LogFile = Join-Path $OutputDirectory ('{0}_{1:yyyyMMdd_HHmmss}.log' -f $Prefix, (Get-Date))
}

function Connect-Tenant([string[]]$Scopes) {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw 'Module Microsoft.Graph.Authentication is missing: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    if ($ClientId) {
        if (-not $CertificateThumbprint) { throw '-CertificateThumbprint is required together with -ClientId.' }
        Write-Log "Connecting to $TenantId (app-only, client $ClientId)"
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop
    }
    else {
        Write-Log "Connecting to $TenantId (delegated: $($Scopes -join ', '))"
        Connect-MgGraph -TenantId $TenantId -Scopes $Scopes -NoWelcome -ErrorAction Stop
    }
}

function Get-HttpStatus($ErrorRecord) {
    try { $code = [int]$ErrorRecord.Exception.Response.StatusCode; if ($code) { return $code } } catch { }
    if ("$($ErrorRecord.Exception.Message)" -match '\b(429|503|504)\b') { return [int]$Matches[1] }
    return 0
}

function Get-RetryDelay($ErrorRecord, [int]$Attempt) {
    try {
        $ra = $ErrorRecord.Exception.Response.Headers.RetryAfter
        if ($ra -and $ra.Delta) { return [int][math]::Ceiling($ra.Delta.TotalSeconds) }
    } catch { }
    return [int][math]::Min(60, [math]::Pow(2, $Attempt))
}

function Invoke-Graph {
    <#
        Calls Graph. For GET collections it follows @odata.nextLink and emits the items.
        Retries on 429/503/504 and honours Retry-After.
    #>
    param(
        [Parameter(Mandatory)][string]$Uri,
        [ValidateSet('GET', 'POST', 'PATCH', 'DELETE')][string]$Method = 'GET',
        $Body,
        [int]$MaxRetries = 6
    )
    $next = $Uri
    while ($next) {
        $attempt = 0
        while ($true) {
            try {
                $req = @{ Method = $Method; Uri = $next; OutputType = 'PSObject'; ErrorAction = 'Stop' }
                if ($null -ne $Body) {
                    $req.Body = ($Body | ConvertTo-Json -Depth 20)
                    $req.ContentType = 'application/json'
                }
                $resp = Invoke-MgGraphRequest @req
                break
            }
            catch {
                $status = Get-HttpStatus $_
                if ($status -in 429, 503, 504 -and $attempt -lt $MaxRetries) {
                    $attempt++
                    $delay = Get-RetryDelay $_ $attempt
                    Write-Log "  Graph returned $status; retry $attempt in $delay s" 'WARN'
                    Start-Sleep -Seconds $delay
                    continue
                }
                throw
            }
        }
        if ($Method -ne 'GET' -or $null -eq $resp -or -not $resp.PSObject.Properties['value']) { return $resp }
        foreach ($item in $resp.value) { $item }
        $nl = $resp.PSObject.Properties['@odata.nextLink']
        $next = if ($nl) { $nl.Value } else { $null }
    }
}

function Get-TenantLabel {
    try {
        $org = @(Invoke-Graph -Uri 'v1.0/organization?$select=displayName,verifiedDomains')[0]
        $dom = @($org.verifiedDomains | Where-Object { $_.isDefault }) + @($org.verifiedDomains | Where-Object { $_.isInitial })
        return [pscustomobject]@{ Name = $org.displayName; Domain = $dom[0].name }
    }
    catch {
        Write-Log 'Could not read organization (Organization.Read.All); using the tenant ID as label.' 'WARN'
        return [pscustomobject]@{ Name = $TenantId; Domain = $TenantId }
    }
}

function Export-Report($Rows, [string]$Name) {
    $path = Join-Path $OutputDirectory ('{0}_{1}_{2:yyyyMMdd}.csv' -f $Name, $script:Tenant.Domain, (Get-Date))
    $list = @($Rows | Where-Object { $null -ne $_ })
    $list | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8 -Delimiter $Delimiter -WhatIf:$false
    Write-Log "Exported $($list.Count) row(s) to $path"
    return $path
}

#endregion

#region Script-specific helpers -----------------------------------------------------------------

function Get-IpNamedLocations {
    @(Invoke-Graph -Uri 'v1.0/identity/conditionalAccess/namedLocations') |
        Where-Object { $_.'@odata.type' -eq '#microsoft.graph.ipNamedLocation' }
}

function Get-Evaluation {
    param($Ips, $Locations, $Policies)

    foreach ($ip in $Ips) {
        $hits = New-Object System.Collections.Generic.List[object]
        foreach ($loc in $Locations) {
            foreach ($r in $loc.ipRanges) {
                $contains = $false
                try { $contains = Test-CidrContains -Network $r.cidrAddress -Address $ip } catch { }
                if ($contains) { $hits.Add($loc); break }
            }
        }
        $trusted = @($hits | Where-Object { $_.isTrusted }).Count -gt 0

        $include = New-Object System.Collections.Generic.List[string]
        $exclude = New-Object System.Collections.Generic.List[string]
        foreach ($p in $Policies) {
            $cond = $p.conditions.locations
            if (-not $cond) { continue }
            $label = '{0} [{1}]' -f $p.displayName, $p.state
            $inc = @($cond.includeLocations); $exc = @($cond.excludeLocations)
            foreach ($h in $hits) {
                if ($inc -contains $h.id) { $include.Add("$label via '$($h.displayName)'") }
                if ($exc -contains $h.id) { $exclude.Add("$label via '$($h.displayName)'") }
            }
            if ($trusted -and $inc -contains 'AllTrusted') { $include.Add("$label via AllTrusted") }
            if ($trusted -and $exc -contains 'AllTrusted') { $exclude.Add("$label via AllTrusted") }
        }

        [pscustomobject][ordered]@{
            IpAddress         = $ip
            InNamedLocation   = ($hits.Count -gt 0)
            Trusted           = $trusted
            MatchedLocations  = (($hits | ForEach-Object { '{0} ({1})' -f $_.displayName, $(if ($_.isTrusted) { 'trusted' } else { 'not trusted' }) }) -join '; ')
            MatchedIds        = (($hits | ForEach-Object { $_.id }) -join ';')
            PoliciesIncluding = ($include -join '; ')
            PoliciesExcluding = ($exclude -join '; ')
        }
    }
}

function New-RangeObject([string]$Cidr) {
    $p = ConvertTo-IpPrefix $Cidr
    $type = if ($p.Family -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6) { '#microsoft.graph.iPv6CidrRange' } else { '#microsoft.graph.iPv4CidrRange' }
    [ordered]@{ '@odata.type' = $type; cidrAddress = (Get-NetworkCidr $Cidr) }
}

#endregion

#region Main ---------------------------------------------------------------------------------

$exitCode = 2
try {
    # Validate input before connecting
    foreach ($ip in $IpAddress) { $null = ConvertTo-IpPrefix $ip }

    Initialize-Output 'NamedLocations'
    Connect-Tenant -Scopes $Scopes
    $script:Tenant = Get-TenantLabel
    Write-Log "=== Named location check for $($script:Tenant.Name) ($($script:Tenant.Domain)): $($IpAddress -join ', ') ==="

    $locations = @(Get-IpNamedLocations)
    $policies  = @(Invoke-Graph -Uri 'v1.0/identity/conditionalAccess/policies')
    Write-Log "IP named locations: $($locations.Count) ($(@($locations | Where-Object { $_.isTrusted }).Count) trusted); CA policies: $($policies.Count)"

    $result = @(Get-Evaluation -Ips $IpAddress -Locations $locations -Policies $policies)
    foreach ($r in $result) {
        $lvl = if ($r.Trusted) { 'INFO' } else { 'WARN' }
        Write-Log ("{0}: trusted={1}; locations: {2}" -f $r.IpAddress, $r.Trusted, $(if ($r.MatchedLocations) { $r.MatchedLocations } else { '(none)' })) $lvl
        if ($r.PoliciesIncluding) { Write-Log "  Included by: $($r.PoliciesIncluding)" }
        if ($r.PoliciesExcluding) { Write-Log "  Excluded by: $($r.PoliciesExcluding)" }
    }

    # -Fix -------------------------------------------------------------------------------------
    $notTrusted = @($result | Where-Object { -not $_.Trusted })
    if ($Fix -and $notTrusted.Count -gt 0) {
        $changed = $false

        if ($LocationName) {
            $target  = @($locations | Where-Object { $_.displayName -eq $LocationName })[0]
            $newCidr = @($notTrusted | ForEach-Object { Get-NetworkCidr $_.IpAddress } | Select-Object -Unique)

            if ($target) {
                $ranges = New-Object System.Collections.Generic.List[object]
                foreach ($r in $target.ipRanges) { $ranges.Add([ordered]@{ '@odata.type' = $r.'@odata.type'; cidrAddress = $r.cidrAddress }) }
                foreach ($c in $newCidr) {
                    if (@($ranges | Where-Object { $_.cidrAddress -eq $c }).Count -eq 0) { $ranges.Add((New-RangeObject $c)) }
                }
                $body = [ordered]@{
                    '@odata.type' = '#microsoft.graph.ipNamedLocation'
                    isTrusted     = $true
                    ipRanges      = [object[]]$ranges.ToArray()
                }
                if ($PSCmdlet.ShouldProcess("Named location '$LocationName'", "Add $($newCidr -join ', ') and mark trusted")) {
                    $null = Invoke-Graph -Method PATCH -Uri "v1.0/identity/conditionalAccess/namedLocations/$($target.id)" -Body $body
                    Write-Log "Named location '$LocationName' updated: added $($newCidr -join ', '), isTrusted = true"
                    $changed = $true
                }
            }
            else {
                $body = [ordered]@{
                    '@odata.type' = '#microsoft.graph.ipNamedLocation'
                    displayName   = $LocationName
                    isTrusted     = $true
                    ipRanges      = [object[]]@($newCidr | ForEach-Object { New-RangeObject $_ })
                }
                if ($PSCmdlet.ShouldProcess("Named location '$LocationName'", "Create as trusted with $($newCidr -join ', ')")) {
                    $null = Invoke-Graph -Method POST -Uri 'v1.0/identity/conditionalAccess/namedLocations' -Body $body
                    Write-Log "Named location '$LocationName' created (trusted) with $($newCidr -join ', ')"
                    $changed = $true
                }
            }
        }
        else {
            $toTrust = @($notTrusted | Where-Object { $_.MatchedIds } | ForEach-Object { $_.MatchedIds -split ';' } | Select-Object -Unique)
            foreach ($id in $toTrust) {
                $loc = @($locations | Where-Object { $_.id -eq $id })[0]
                $body = [ordered]@{ '@odata.type' = '#microsoft.graph.ipNamedLocation'; isTrusted = $true }
                if ($PSCmdlet.ShouldProcess("Named location '$($loc.displayName)' (all its ranges)", 'Mark as trusted')) {
                    $null = Invoke-Graph -Method PATCH -Uri "v1.0/identity/conditionalAccess/namedLocations/$id" -Body $body
                    Write-Log "Named location '$($loc.displayName)' marked as trusted"
                    $changed = $true
                }
            }
            $orphans = @($notTrusted | Where-Object { -not $_.InNamedLocation })
            if ($orphans.Count -gt 0) {
                Write-Log ("Not in any named location: {0}. Use -LocationName to add them." -f (($orphans | ForEach-Object IpAddress) -join ', ')) 'WARN'
            }
        }

        if ($changed) {
            Write-Log 'Re-evaluating...'
            $locations = @(Get-IpNamedLocations)
            $result = @(Get-Evaluation -Ips $IpAddress -Locations $locations -Policies $policies)
        }
    }

    $null = Export-Report $result 'NamedLocations'

    Write-Host ''
    $result | Format-Table IpAddress, Trusted, MatchedLocations -AutoSize | Out-String | Write-Host
    $bad = @($result | Where-Object { -not $_.Trusted })
    Write-Log ("Summary: {0} of {1} IP(s) in a trusted named location." -f ($result.Count - $bad.Count), $result.Count)

    $exitCode = if ($bad.Count -gt 0) { 1 } else { 0 }
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    $exitCode = 2
}
finally {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
}
exit $exitCode

#endregion
