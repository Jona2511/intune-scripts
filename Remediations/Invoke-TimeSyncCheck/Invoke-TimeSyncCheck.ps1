#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Checks the Windows Time service and clock drift, and repairs them if needed.

.DESCRIPTION
    Checks:
      - Whether the Windows Time service (W32Time) is disabled or not running.
      - Which time source the client uses. 'Local CMOS Clock' or 'Free-running System
        Clock' means it is not syncing with anything.
      - The clock offset compared with a reference server, measured with
        'w32tm /stripchart'. The median of several samples is used.

    Repair (unless -DetectOnly):
      - Disabled service -> startup type set to Automatic.
      - Service stopped or drift above -MaxDriftSeconds -> service restarted,
        then 'w32tm /resync /force'. If that fails: 'w32tm /config /update' and
        'w32tm /resync /rediscover'.
      - With -FixSource on a non-domain machine without a valid source: sets -NtpServer
        as the NTP source. Domain machines always keep the domain hierarchy.
      - Measures the drift again afterwards.

.PARAMETER MaxDriftSeconds
    Maximum allowed offset in seconds (default 5). Kerberos fails above 300.

.PARAMETER ReferenceServer
    Server to measure the drift against. Default: the current time source, or -NtpServer
    if there is no valid source.

.PARAMETER NtpServer
    NTP server used as a fallback reference and for -FixSource.

.PARAMETER FixSource
    Configures -NtpServer as the time source on non-domain machines without a valid source.

.PARAMETER DetectOnly
    Only report. Exit code 1 if something is wrong.

.EXAMPLE
    .\Invoke-TimeSyncCheck.ps1 -DetectOnly

.EXAMPLE
    .\Invoke-TimeSyncCheck.ps1 -MaxDriftSeconds 2 -FixSource -NtpServer 'nl.pool.ntp.org'

.NOTES
    Exit codes: 0 = in order (afterwards), 1 = problem remains, 2 = fatal error.
    Log: %ProgramData%\TimeSyncCheck\
    'w32tm /stripchart' needs outbound UDP 123 to the reference server.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateRange(0.1, 3600)]
    [double]$MaxDriftSeconds = 5,

    [string]$ReferenceServer,

    [string]$NtpServer = 'time.windows.com',

    [switch]$FixSource,
    [switch]$DetectOnly,

    [ValidateRange(1, 20)]
    [int]$Samples = 3,

    [string]$LogDirectory = (Join-Path $env:ProgramData 'TimeSyncCheck')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$W32tm = Join-Path $env:windir 'System32\w32tm.exe'

#region Helpers ------------------------------------------------------------------------------

$script:LogFile = $null

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
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

function Invoke-W32tm {
    param([Parameter(Mandatory)][string[]]$Arguments)
    $out = & $W32tm @Arguments 2>&1 | ForEach-Object { "$_" }
    [pscustomobject]@{
        ExitCode = $LASTEXITCODE
        Output   = ($out -join "`n").Trim()
    }
}

function Test-DomainJoined {
    (Get-CimInstance -ClassName Win32_ComputerSystem).PartOfDomain
}

function Get-TimeSource {
    $r = Invoke-W32tm -Arguments '/query', '/source'
    if ($r.ExitCode -ne 0) { return $null }
    # e.g. "time.windows.com,0x9" or "DC01.contoso.local"
    return ($r.Output -split '[,\s]')[0].Trim()
}

function Test-ValidSource([string]$Source) {
    if ([string]::IsNullOrWhiteSpace($Source)) { return $false }
    # 'Local CMOS Clock', 'Free-running System Clock' and the VM IC provider are not network sources
    return -not ($Source -match 'CMOS|Free-running|^Local$|^Free$|VM IC')
}

function Get-ClockOffset {
    <#
        Returns the median offset in seconds compared with $Server, or $null if it
        cannot be measured. Handles both '.' and ',' as decimal separator (Dutch locale).
    #>
    param([Parameter(Mandatory)][string]$Server)

    $r = Invoke-W32tm -Arguments '/stripchart', "/computer:$Server", '/dataonly', "/samples:$Samples"
    $values = foreach ($m in [regex]::Matches($r.Output, '([+-]\d+[.,]\d+)s')) {
        [double]::Parse($m.Groups[1].Value.Replace(',', '.'), [Globalization.CultureInfo]::InvariantCulture)
    }
    $values = @($values)
    if ($values.Count -eq 0) {
        $firstErr = ($r.Output -split "`n" | Where-Object { $_ -match 'error|fout|0x8' } | Select-Object -First 1)
        Write-Log "  Could not measure offset against ${Server}: $firstErr" 'WARN'
        return $null
    }
    $sorted = $values | Sort-Object
    return [double]$sorted[[int][math]::Floor(($sorted.Count - 1) / 2)]
}

function Measure-Drift {
    param([string]$Server)
    $offset = Get-ClockOffset -Server $Server
    if ($null -eq $offset) { return $null }
    Write-Log ("  Offset against {0}: {1:+0.000;-0.000} s (limit {2} s)" -f $Server, $offset, $MaxDriftSeconds)
    return $offset
}

#endregion

#region Main ---------------------------------------------------------------------------------

try {
    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force -WhatIf:$false | Out-Null
    }
    $script:LogFile = Join-Path $LogDirectory ('TimeSync_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
    Get-ChildItem -LiteralPath $LogDirectory -Filter 'TimeSync_*.log' |
        Sort-Object LastWriteTime -Descending | Select-Object -Skip 30 |
        Remove-Item -Force -WhatIf:$false -ErrorAction SilentlyContinue

    $mode = if ($DetectOnly) { 'DetectOnly' } elseif ($WhatIfPreference) { 'WhatIf' } else { 'Remediate' }
    $isDomain = Test-DomainJoined
    Write-Log "=== Time sync check started on $env:COMPUTERNAME ($mode, domain-joined: $isDomain) ==="

    $problems = New-Object System.Collections.Generic.List[string]

    # 1. Service --------------------------------------------------------------------------
    $svc = Get-CimInstance -ClassName Win32_Service -Filter "Name='W32Time'"
    if (-not $svc) { throw 'Windows Time service (W32Time) is not installed.' }
    Write-Log "W32Time: state $($svc.State), startup type $($svc.StartMode)"

    $serviceDisabled = $svc.StartMode -eq 'Disabled'
    $serviceStopped  = $svc.State -ne 'Running'
    if ($serviceDisabled) { $problems.Add('service disabled') }
    if ($serviceStopped)  { $problems.Add('service not running') }

    # 2. Time source ----------------------------------------------------------------------
    $source = if ($serviceStopped) { $null } else { Get-TimeSource }
    $sourceValid = Test-ValidSource $source
    Write-Log ("Current time source: {0}" -f $(if ($source) { $source } else { '(unknown)' }))
    if (-not $serviceStopped -and -not $sourceValid) {
        Write-Log '  No valid network time source.' 'WARN'
        $problems.Add('no valid time source')
    }

    # 3. Drift ----------------------------------------------------------------------------
    $reference = if ($ReferenceServer) { $ReferenceServer }
                 elseif ($sourceValid) { $source }
                 else { $NtpServer }

    $offset = Measure-Drift -Server $reference
    if ($null -eq $offset) {
        $problems.Add('drift not measurable')
    }
    elseif ([math]::Abs($offset) -gt $MaxDriftSeconds) {
        Write-Log '  Clock drift is above the limit.' 'WARN'
        $problems.Add(('drift {0:N1}s' -f $offset))
    }

    if ($problems.Count -eq 0) {
        Write-Log 'Everything in order.'
        Write-Output ('COMPLIANT: W32Time running, source {0}, offset {1:N3}s' -f $source, $offset)
        exit 0
    }

    Write-Log ("Problems found: {0}" -f ($problems -join ', ')) 'WARN'
    if ($DetectOnly) {
        Write-Output ('NON-COMPLIANT: ' + ($problems -join ', '))
        exit 1
    }

    # 4. Repair ---------------------------------------------------------------------------
    if ($serviceDisabled -and $PSCmdlet.ShouldProcess('W32Time', 'Set startup type to Automatic')) {
        Set-Service -Name W32Time -StartupType Automatic
        Write-Log 'W32Time startup type set to Automatic.'
    }

    if (-not $sourceValid -and -not $serviceStopped) {
        if ($isDomain) {
            Write-Log 'Domain machine without a valid source: resetting to domain hierarchy.' 'WARN'
            if ($PSCmdlet.ShouldProcess('W32Time', 'Configure /syncfromflags:domhier')) {
                $null = Invoke-W32tm -Arguments '/config', '/syncfromflags:domhier', '/update'
            }
        }
        elseif ($FixSource) {
            if ($PSCmdlet.ShouldProcess('W32Time', "Set NTP source to $NtpServer")) {
                $r = Invoke-W32tm -Arguments '/config', "/manualpeerlist:$NtpServer,0x9", '/syncfromflags:manual', '/reliable:no', '/update'
                Write-Log "NTP source set to $NtpServer (exit code $($r.ExitCode))."
            }
        }
        else {
            Write-Log 'No valid source; use -FixSource to configure one.' 'WARN'
        }
    }

    if ($PSCmdlet.ShouldProcess('W32Time', 'Restart service and resync')) {
        Write-Log 'Restarting W32Time...'
        if ((Get-Service W32Time).Status -eq 'Running') {
            Restart-Service -Name W32Time -Force
        }
        else {
            Start-Service -Name W32Time
        }
        (Get-Service W32Time).WaitForStatus('Running', [TimeSpan]::FromSeconds(30))

        $r = Invoke-W32tm -Arguments '/resync', '/force'
        if ($r.ExitCode -ne 0) {
            Write-Log "Resync failed ($($r.Output)); retrying with /rediscover." 'WARN'
            $null = Invoke-W32tm -Arguments '/config', '/update'
            $r = Invoke-W32tm -Arguments '/resync', '/rediscover'
        }
        if ($r.ExitCode -eq 0) { Write-Log 'Resync succeeded.' }
        else { Write-Log "Resync failed: $($r.Output)" 'ERROR' }
    }

    if ($WhatIfPreference) {
        Write-Output ('NON-COMPLIANT (WhatIf, nothing changed): ' + ($problems -join ', '))
        exit 1
    }

    # 5. Verify ---------------------------------------------------------------------------
    Write-Log 'Verifying...'
    $source      = Get-TimeSource
    $sourceValid = Test-ValidSource $source
    $reference   = if ($ReferenceServer) { $ReferenceServer } elseif ($sourceValid) { $source } else { $NtpServer }
    $offset      = Measure-Drift -Server $reference
    $running     = (Get-Service W32Time).Status -eq 'Running'

    $remaining = New-Object System.Collections.Generic.List[string]
    if (-not $running)                                   { $remaining.Add('service not running') }
    if (-not $sourceValid)                               { $remaining.Add("invalid source '$source'") }
    if ($null -eq $offset)                               { $remaining.Add('drift not measurable') }
    elseif ([math]::Abs($offset) -gt $MaxDriftSeconds)   { $remaining.Add(('drift {0:N1}s' -f $offset)) }

    if ($remaining.Count -gt 0) {
        Write-Log ("Still not in order: {0}" -f ($remaining -join ', ')) 'ERROR'
        Write-Output ('NON-COMPLIANT after repair: ' + ($remaining -join ', '))
        exit 1
    }

    Write-Log ('Repaired. Source {0}, offset {1:N3}s.' -f $source, $offset)
    Write-Output ('COMPLIANT after repair: source {0}, offset {1:N3}s' -f $source, $offset)
    exit 0
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 2
}

#endregion
