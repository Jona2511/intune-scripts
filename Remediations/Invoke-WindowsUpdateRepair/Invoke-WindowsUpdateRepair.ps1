#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Detects a stuck Windows Update and resets the update components.

.DESCRIPTION
    Detection:
      - Date of the last successfully installed update, taken from the Windows Update
        history (Microsoft.Update.Session). By default only updates with a KB number count;
        Defender definitions, MSRT and Store app updates are ignored because they install
        often and would hide a stuck cumulative update.
        Fallback when the history is empty: Get-HotFix.
      - Whether the wuauserv service is disabled.
      - Whether a reboot is pending. In that case a reset is pointless: the script reports
        it and skips the reset (unless -IgnorePendingReboot).

    Reset (when the last update is older than -MaxDaysSinceUpdate):
      1. Stops wuauserv, bits, cryptsvc, msiserver and usosvc.
      2. Deletes the BITS queue (qmgr*.dat).
      3. Renames SoftwareDistribution and catroot2 to *.bak_<timestamp>
         (older .bak folders from this script are removed).
      4. Sets wuauserv back to Manual if it was disabled.
      5. Starts the services again and triggers a new scan.
      Optional with -RunDism: DISM /RestoreHealth and SFC /scannow (takes a long time).

    Resetting clears the visible update history in Settings. To avoid resetting every
    run, the script waits at least -MinDaysBetweenResets after a previous reset.

.PARAMETER MaxDaysSinceUpdate
    The reset runs if the last successful update is older than this (default 35).

.PARAMETER MinDaysBetweenResets
    Minimum number of days between two resets (default 7).

.PARAMETER CountAllUpdates
    Count every successful update, including Defender definitions, Store apps and drivers.
    By default only updates with a KB number count (Windows, .NET, Office), excluding
    Defender/MSRT updates.

.PARAMETER IgnorePendingReboot
    Resets even when a reboot is pending.

.PARAMETER RunDism
    Also runs DISM /RestoreHealth and SFC /scannow after the reset.

.PARAMETER Force
    Resets regardless of the update age and the time since the last reset.

.PARAMETER DetectOnly
    Only report. Exit code 1 if Windows Update looks stuck.

.EXAMPLE
    .\Invoke-WindowsUpdateRepair.ps1 -DetectOnly

.EXAMPLE
    .\Invoke-WindowsUpdateRepair.ps1 -MaxDaysSinceUpdate 45 -RunDism

.NOTES
    Exit codes: 0 = in order or repaired, 1 = stuck / reboot needed / reset failed,
    2 = fatal error.
    Log: %ProgramData%\WindowsUpdateRepair\
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateRange(1, 3650)]
    [int]$MaxDaysSinceUpdate = 35,

    [ValidateRange(0, 365)]
    [int]$MinDaysBetweenResets = 7,

    [switch]$CountAllUpdates,
    [switch]$IgnorePendingReboot,
    [switch]$RunDism,
    [switch]$Force,
    [switch]$DetectOnly,

    [string]$LogDirectory = (Join-Path $env:ProgramData 'WindowsUpdateRepair')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$Services = 'wuauserv', 'bits', 'cryptsvc', 'msiserver', 'usosvc'

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

function Get-LastSuccessfulUpdate {
    <# Returns [pscustomobject]@{ Date; Title; Source } or $null. #>
    # Defender definitions/platform, MSRT and Defender for Endpoint update far more often than
    # quality updates and would hide a stuck Windows Update.
    $definitionPattern = 'Security Intelligence|Definition Update|Defender Antivirus|KB2267602|KB2461484|KB4052623|KB915597|KB890830|KB5005292'

    try {
        $session  = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $count    = $searcher.GetTotalHistoryCount()
        if ($count -gt 0) {
            $latest = $searcher.QueryHistory(0, $count) | Where-Object {
                $_.Operation -eq 1 -and                    # 1 = installation
                $_.ResultCode -in 2, 3 -and                # 2 = succeeded, 3 = succeeded with errors
                ($CountAllUpdates -or
                    # Only Windows/Office/.NET updates with a KB number; Store apps have none
                    ($_.Title -match 'KB\d{6,}' -and $_.Title -notmatch $definitionPattern))
            } | Sort-Object Date -Descending | Select-Object -First 1

            if ($latest) {
                return [pscustomobject]@{
                    Date   = $latest.Date.ToLocalTime()
                    Title  = $latest.Title
                    Source = 'Windows Update history'
                }
            }
        }
    }
    catch { Write-Log "Could not read update history: $($_.Exception.Message)" 'WARN' }

    # Fallback: CBS hotfixes (e.g. after a previous reset that cleared the history)
    $hotfix = Get-HotFix -ErrorAction SilentlyContinue |
        Where-Object { $_.InstalledOn } |
        Sort-Object InstalledOn -Descending | Select-Object -First 1
    if ($hotfix) {
        return [pscustomobject]@{
            Date   = [datetime]$hotfix.InstalledOn
            Title  = "$($hotfix.HotFixID) ($($hotfix.Description))"
            Source = 'Get-HotFix'
        }
    }
    return $null
}

function Test-PendingReboot {
    $keys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
    )
    foreach ($k in $keys) { if (Test-Path -LiteralPath $k) { return $true } }
    return $false
}

function Stop-UpdateServices {
    foreach ($name in $Services) {
        $svc = Get-Service -Name $name -ErrorAction SilentlyContinue
        if (-not $svc -or $svc.Status -eq 'Stopped') { continue }
        Write-Log "  Stopping $name"
        for ($i = 1; $i -le 3; $i++) {
            try {
                Stop-Service -Name $name -Force -ErrorAction Stop
                $svc.WaitForStatus('Stopped', [TimeSpan]::FromSeconds(30))
                break
            }
            catch {
                if ($i -eq 3) { Write-Log "  Could not stop ${name}: $($_.Exception.Message)" 'WARN' }
                Start-Sleep -Seconds 3
            }
        }
    }
}

function Start-UpdateServices {
    foreach ($name in 'cryptsvc', 'bits', 'wuauserv', 'usosvc') {
        $svc = Get-Service -Name $name -ErrorAction SilentlyContinue
        if (-not $svc -or $svc.Status -eq 'Running') { continue }
        try {
            Start-Service -Name $name -ErrorAction Stop
            Write-Log "  Started $name"
        }
        catch { Write-Log "  Could not start ${name}: $($_.Exception.Message)" 'WARN' }
    }
}

function Rename-WithBackup {
    <# Renames a folder to <name>.bak_<timestamp> and removes older backups of it. #>
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Stamp)

    $parent = Split-Path -Path $Path -Parent
    $leaf   = Split-Path -Path $Path -Leaf

    Get-ChildItem -LiteralPath $parent -Directory -Filter "$leaf.bak_*" -ErrorAction SilentlyContinue |
        ForEach-Object {
            Write-Log "  Removing old backup $($_.FullName)"
            Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
        }

    if (-not (Test-Path -LiteralPath $Path)) { return $true }

    $newName = "$leaf.bak_$Stamp"
    for ($i = 1; $i -le 3; $i++) {
        try {
            Rename-Item -LiteralPath $Path -NewName $newName -ErrorAction Stop
            Write-Log "  $Path -> $newName"
            return $true
        }
        catch {
            if ($i -eq 3) {
                Write-Log "  Rename of $Path failed: $($_.Exception.Message)" 'ERROR'
                return $false
            }
            # A service (often cryptsvc) may have restarted itself; stop it again
            Stop-UpdateServices
            Start-Sleep -Seconds 3
        }
    }
}

function Invoke-UpdateReset {
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $ok = $true

    Write-Log 'Stopping update services...'
    Stop-UpdateServices

    try {
        Write-Log 'Clearing BITS queue...'
        $downloader = Join-Path $env:ProgramData 'Microsoft\Network\Downloader'
        Get-ChildItem -LiteralPath $downloader -Filter 'qmgr*.dat' -Force -ErrorAction SilentlyContinue |
            Remove-Item -Force -ErrorAction SilentlyContinue

        Write-Log 'Renaming SoftwareDistribution and catroot2...'
        $ok = (Rename-WithBackup -Path (Join-Path $env:windir 'SoftwareDistribution') -Stamp $stamp) -and $ok
        $ok = (Rename-WithBackup -Path (Join-Path $env:windir 'System32\catroot2') -Stamp $stamp) -and $ok

        $wu = Get-CimInstance -ClassName Win32_Service -Filter "Name='wuauserv'"
        if ($wu.StartMode -eq 'Disabled') {
            Set-Service -Name wuauserv -StartupType Manual
            Write-Log 'wuauserv startup type set to Manual (default).'
        }
    }
    finally {
        Write-Log 'Starting update services...'
        Start-UpdateServices
    }

    if ($RunDism) {
        Write-Log 'DISM /RestoreHealth (may take a long time)...'
        & "$env:windir\System32\dism.exe" /Online /Cleanup-Image /RestoreHealth /NoRestart | Out-Null
        Write-Log "  DISM exit code: $LASTEXITCODE" $(if ($LASTEXITCODE -eq 0) { 'INFO' } else { 'WARN' })

        Write-Log 'SFC /scannow...'
        & "$env:windir\System32\sfc.exe" /scannow | Out-Null
        Write-Log "  SFC exit code: $LASTEXITCODE"
    }

    Write-Log 'Triggering an update scan...'
    try {
        $uso = Join-Path $env:windir 'System32\UsoClient.exe'
        if (Test-Path -LiteralPath $uso) { & $uso StartScan | Out-Null }
        else { (New-Object -ComObject Microsoft.Update.AutoUpdate).DetectNow() }
    }
    catch { Write-Log "  Starting the scan failed: $($_.Exception.Message)" 'WARN' }

    return $ok
}

#endregion

#region Main ---------------------------------------------------------------------------------

try {
    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force -WhatIf:$false | Out-Null
    }
    $script:LogFile = Join-Path $LogDirectory ('WURepair_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
    Get-ChildItem -LiteralPath $LogDirectory -Filter 'WURepair_*.log' |
        Sort-Object LastWriteTime -Descending | Select-Object -Skip 30 |
        Remove-Item -Force -WhatIf:$false -ErrorAction SilentlyContinue
    $markerFile = Join-Path $LogDirectory 'last-reset.txt'

    $mode = if ($DetectOnly) { 'DetectOnly' } elseif ($WhatIfPreference) { 'WhatIf' } else { 'Remediate' }
    Write-Log "=== Windows Update check started on $env:COMPUTERNAME ($mode) ==="

    # Informational: WSUS or Windows Update for Business?
    $wuPolicy = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' -ErrorAction SilentlyContinue
    if ($wuPolicy -and $wuPolicy.PSObject.Properties['WUServer']) {
        Write-Log "WSUS server configured: $($wuPolicy.WUServer)"
    }

    # 1. Detection -------------------------------------------------------------------------
    $last = Get-LastSuccessfulUpdate
    if ($last) {
        $age = [math]::Floor(((Get-Date) - $last.Date).TotalDays)
        Write-Log ("Last successful update: {0:yyyy-MM-dd} ({1} days ago) - {2} [{3}]" -f `
            $last.Date, $age, $last.Title, $last.Source)
    }
    else {
        $age = [int]::MaxValue
        Write-Log 'No successful update found.' 'WARN'
    }

    $wuSvc        = Get-CimInstance -ClassName Win32_Service -Filter "Name='wuauserv'"
    $wuDisabled   = $wuSvc.StartMode -eq 'Disabled'
    $pendingBoot  = Test-PendingReboot
    $tooOld       = $age -gt $MaxDaysSinceUpdate

    if ($wuDisabled)  { Write-Log 'wuauserv is disabled.' 'WARN' }
    if ($pendingBoot) { Write-Log 'A reboot is pending.' 'WARN' }

    if (-not $tooOld -and -not $wuDisabled -and -not $Force) {
        Write-Log "Windows Update is fine (limit $MaxDaysSinceUpdate days)."
        Write-Output ("COMPLIANT: last update {0} days ago" -f $age)
        exit 0
    }

    $reason = @()
    if ($tooOld)     { $reason += $(if ($last) { "last update $age days ago" } else { 'no successful update found' }) }
    if ($wuDisabled) { $reason += 'wuauserv disabled' }
    if ($pendingBoot) { $reason += 'reboot pending' }
    $reasonText = $reason -join ', '

    if ($DetectOnly) {
        Write-Output "NON-COMPLIANT: $reasonText"
        exit 1
    }

    # 2. Guards --------------------------------------------------------------------------
    if ($pendingBoot -and -not $IgnorePendingReboot -and -not $Force) {
        Write-Log 'Reset skipped: reboot first, updates may be waiting for it.' 'WARN'
        Write-Output "NON-COMPLIANT: $reasonText - reboot required, reset skipped"
        exit 1
    }

    if ((Test-Path -LiteralPath $markerFile) -and -not $Force) {
        $lastReset = [datetime]::ParseExact((Get-Content -LiteralPath $markerFile -TotalCount 1).Trim(),
            'yyyy-MM-dd HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
        $sinceReset = ((Get-Date) - $lastReset).TotalDays
        if ($sinceReset -lt $MinDaysBetweenResets) {
            Write-Log ("Previous reset was {0:N1} days ago (minimum {1}); giving Windows Update time." -f `
                $sinceReset, $MinDaysBetweenResets) 'WARN'
            Write-Output ("NON-COMPLIANT: $reasonText - reset already done on {0:yyyy-MM-dd}" -f $lastReset)
            exit 1
        }
    }

    # 3. Reset ---------------------------------------------------------------------------
    if (-not $PSCmdlet.ShouldProcess($env:COMPUTERNAME, 'Reset Windows Update components')) {
        Write-Output "NON-COMPLIANT (WhatIf, nothing changed): $reasonText"
        exit 1
    }

    Write-Log "Resetting Windows Update components ($reasonText)..."
    $ok = Invoke-UpdateReset
    Set-Content -LiteralPath $markerFile -Value (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') -Encoding UTF8

    $notRunning = @('wuauserv', 'bits', 'cryptsvc') | Where-Object { (Get-Service $_).Status -ne 'Running' }
    if ($notRunning) {
        # wuauserv/bits are trigger-start and may stop again by themselves; that is fine as long as they are not disabled
        Write-Log ("Not running after reset (may be normal for trigger-start services): {0}" -f ($notRunning -join ', '))
    }

    if (-not $ok) {
        Write-Log 'Reset partially failed; see log.' 'ERROR'
        Write-Output "NON-COMPLIANT: reset partially failed ($reasonText)"
        exit 1
    }

    Write-Log 'Reset complete. Windows Update is scanning again; installation follows in the background.'
    Write-Output "REMEDIATED: Windows Update components reset ($reasonText)"
    exit 0
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 2
}

#endregion
