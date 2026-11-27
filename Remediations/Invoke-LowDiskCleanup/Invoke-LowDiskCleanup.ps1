#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Cleans up C: when free space drops below a threshold.

.DESCRIPTION
    Checks free space on C:. If it is below -MinFreeGB, the script runs these steps,
    from least to most invasive, and stops once enough space is free again:

      1. Temp folders     : C:\Windows\Temp and every user's AppData\Local\Temp
                            (only files older than -TempFileAgeDays).
      2. Windows Update   : SoftwareDistribution\Download and the Delivery Optimization cache.
                            wuauserv and BITS are stopped and restarted afterwards.
      3. Old user profiles: local profiles not used for -ProfileAgeDays days, removed through
                            Win32_UserProfile (folder and registry entry). Profiles that are
                            loaded, special, the current user or in -ExcludeProfiles are skipped.

    Supports -WhatIf: shows what would be deleted without changing anything.

.PARAMETER MinFreeGB
    Threshold in GB. The cleanup runs only below this value (unless -Force).

.PARAMETER ProfileAgeDays
    Profiles not used for this many days are considered old.

.PARAMETER TempFileAgeDays
    Only temp files older than this many days are deleted (0 = everything).

.PARAMETER ExcludeProfiles
    Profile folder names (wildcards allowed) that are never removed.

.PARAMETER Force
    Runs all steps regardless of free space.

.EXAMPLE
    .\Invoke-LowDiskCleanup.ps1 -MinFreeGB 15 -WhatIf
    Shows what would be cleaned up if C: has less than 15 GB free.

.EXAMPLE
    .\Invoke-LowDiskCleanup.ps1 -MinFreeGB 20 -ProfileAgeDays 60 -ExcludeProfiles 'admin*','svc_*'

.NOTES
    Exit codes: 0 = enough free space (afterwards), 1 = still below threshold, 2 = fatal error.
    Log: %ProgramData%\DiskCleanup\
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [ValidateRange(1, 100000)]
    [int]$MinFreeGB = 20,

    [ValidateRange(1, 3650)]
    [int]$ProfileAgeDays = 90,

    [ValidateRange(0, 3650)]
    [int]$TempFileAgeDays = 2,

    [string[]]$ExcludeProfiles = @('Administrator', 'Default*', 'Public', 'defaultuser*'),

    [switch]$SkipTemp,
    [switch]$SkipWindowsUpdate,
    [switch]$SkipProfiles,
    [switch]$Force,

    [string]$LogDirectory = (Join-Path $env:ProgramData 'DiskCleanup')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Helpers ------------------------------------------------------------------------------

$script:LogFile = $null

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    switch ($Level) {
        'WARN'  { Write-Warning $Message }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }
    if ($script:LogFile) {
        # Always write the log, even under -WhatIf
        Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -WhatIf:$false
    }
}

function Get-FreeSpaceGB {
    $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
    [math]::Round($disk.FreeSpace / 1GB, 2)
}

function Format-Size([long]$Bytes) {
    if ($Bytes -ge 1GB) { '{0:N2} GB' -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { '{0:N1} MB' -f ($Bytes / 1MB) }
    else { '{0:N0} KB' -f ($Bytes / 1KB) }
}

function Test-BelowThreshold {
    $free = Get-FreeSpaceGB
    Write-Log ("Free space on C: {0} GB (threshold {1} GB)" -f $free, $MinFreeGB)
    return ($free -lt $MinFreeGB)
}

function Clear-FolderContent {
    <#
        Deletes files (older than OlderThanDays) under Path and then removes empty folders.
        Does not follow junctions/symlinks. Locked files are skipped.
        Returns the number of bytes freed (or that would be freed under -WhatIf).
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$OlderThanDays = 0
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return 0L }

    $cutoff  = (Get-Date).AddDays(-$OlderThanDays)
    $execute = $PSCmdlet.ShouldProcess($Path, "Delete files older than $OlderThanDays day(s)")

    $bytes   = 0L
    $files   = 0
    $failed  = 0
    $dirs    = New-Object System.Collections.Generic.List[System.IO.DirectoryInfo]
    $stack   = New-Object System.Collections.Generic.Stack[System.IO.DirectoryInfo]
    $stack.Push([System.IO.DirectoryInfo]::new($Path))

    while ($stack.Count -gt 0) {
        $dir = $stack.Pop()
        try   { $entries = $dir.GetFileSystemInfos() }
        catch { $failed++; continue }

        foreach ($entry in $entries) {
            # Never follow or delete junctions / symlinks
            if ($entry.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { continue }

            if ($entry -is [System.IO.DirectoryInfo]) {
                $stack.Push($entry)
                $dirs.Add($entry)
                continue
            }

            if ($entry.LastWriteTime -ge $cutoff) { continue }

            $size = $entry.Length
            if ($execute) {
                try {
                    if ($entry.Attributes -band [System.IO.FileAttributes]::ReadOnly) {
                        $entry.Attributes = [System.IO.FileAttributes]::Normal
                    }
                    $entry.Delete()
                    $bytes += $size
                    $files++
                }
                catch { $failed++ }
            }
            else {
                $bytes += $size
                $files++
            }
        }
    }

    if ($execute) {
        # Remove empty subfolders, deepest first (Delete() without recursion fails if not empty)
        foreach ($d in ($dirs | Sort-Object { $_.FullName.Length } -Descending)) {
            try {
                $d.Refresh()
                if ($d.Exists -and $d.LastWriteTime -lt $cutoff) { $d.Delete() }
            }
            catch { }
        }
    }

    $verb = if ($execute) { 'Deleted' } else { 'Would delete' }
    Write-Log ("  {0}: {1} {2} file(s), {3}; {4} item(s) skipped (in use/no access)" -f `
        $Path, $verb, $files, (Format-Size $bytes), $failed)
    return $bytes
}

function Get-ProfileLastUse {
    <#
        Win32_UserProfile.LastUseTime is unreliable on Windows 10/11 (it is updated by
        background processes). Use LocalProfileLoad/UnloadTime from the registry instead,
        falling back to NTUSER.DAT.
    #>
    param([string]$Sid, [string]$LocalPath)

    function ConvertTo-UInt32([object]$v) {
        [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$v), 0)
    }

    $key   = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$Sid"
    $props = Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue
    $dates = @()

    foreach ($kind in 'Load', 'Unload') {
        $hiName = "LocalProfile${kind}TimeHigh"
        $loName = "LocalProfile${kind}TimeLow"
        if ($props -and $props.PSObject.Properties[$hiName] -and $props.PSObject.Properties[$loName]) {
            $ft = ([uint64](ConvertTo-UInt32 $props.$hiName) -shl 32) -bor [uint64](ConvertTo-UInt32 $props.$loName)
            if ($ft -gt 0) { $dates += [DateTime]::FromFileTime([int64]$ft) }
        }
    }

    if ($dates.Count -gt 0) { return ($dates | Sort-Object -Descending | Select-Object -First 1) }

    $ntuser = Join-Path $LocalPath 'NTUSER.DAT'
    if (Test-Path -LiteralPath $ntuser) { return (Get-Item -LiteralPath $ntuser -Force).LastWriteTime }
    if (Test-Path -LiteralPath $LocalPath) { return (Get-Item -LiteralPath $LocalPath -Force).LastWriteTime }
    return $null
}

#endregion

#region Steps --------------------------------------------------------------------------------

function Invoke-TempCleanup {
    Write-Log "STEP 1: Temp folders (files older than $TempFileAgeDays day(s))"
    $paths = New-Object System.Collections.Generic.List[string]
    $paths.Add((Join-Path $env:windir 'Temp'))

    Get-CimInstance -ClassName Win32_UserProfile -Filter 'Special=False' |
        Where-Object { $_.LocalPath } |
        ForEach-Object { $paths.Add((Join-Path $_.LocalPath 'AppData\Local\Temp')) }

    $total = 0L
    foreach ($p in ($paths | Select-Object -Unique)) {
        $total += Clear-FolderContent -Path $p -OlderThanDays $TempFileAgeDays
    }
    Write-Log ("STEP 1 done: {0}" -f (Format-Size $total))
}

function Invoke-WindowsUpdateCleanup {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    Write-Log 'STEP 2: Windows Update cache'
    $serviceNames = 'wuauserv', 'bits'
    $wasRunning   = @()

    if ($PSCmdlet.ShouldProcess(($serviceNames -join ', '), 'Stop services temporarily')) {
        foreach ($name in $serviceNames) {
            $svc = Get-Service -Name $name -ErrorAction SilentlyContinue
            if ($svc -and $svc.Status -ne 'Stopped') {
                $wasRunning += $name
                Write-Log "  Stopping service $name"
                Stop-Service -Name $name -Force -ErrorAction Stop
                $svc.WaitForStatus('Stopped', [TimeSpan]::FromSeconds(60))
            }
        }
    }

    try {
        $download = Join-Path $env:windir 'SoftwareDistribution\Download'
        $null = Clear-FolderContent -Path $download -OlderThanDays 0
    }
    finally {
        foreach ($name in $wasRunning) {
            try {
                Write-Log "  Starting service $name"
                Start-Service -Name $name -ErrorAction Stop
            }
            catch { Write-Log "  Could not start service ${name}: $($_.Exception.Message)" 'WARN' }
        }
    }

    if (Get-Command -Name Delete-DeliveryOptimizationCache -ErrorAction SilentlyContinue) {
        if ($PSCmdlet.ShouldProcess('Delivery Optimization cache', 'Clear')) {
            try {
                Delete-DeliveryOptimizationCache -Force -ErrorAction Stop
                Write-Log '  Delivery Optimization cache cleared'
            }
            catch { Write-Log "  Delivery Optimization cache: $($_.Exception.Message)" 'WARN' }
        }
    }
    Write-Log 'STEP 2 done'
}

function Invoke-OldProfileCleanup {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    Write-Log "STEP 3: User profiles not used for more than $ProfileAgeDays day(s)"
    $cutoff     = (Get-Date).AddDays(-$ProfileAgeDays)
    $currentSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value

    $profiles = Get-CimInstance -ClassName Win32_UserProfile -Filter 'Special=False' |
        Where-Object { $_.LocalPath -and $_.SID -like 'S-1-5-21-*' }

    $removed = 0
    foreach ($p in $profiles) {
        $name = Split-Path -Path $p.LocalPath -Leaf

        if ($p.SID -eq $currentSid) { continue }
        if ($p.Loaded) { Write-Log "  Skipped $name (loaded/in use)"; continue }
        if ($ExcludeProfiles | Where-Object { $name -like $_ }) {
            Write-Log "  Skipped $name (excluded)"
            continue
        }

        $lastUse = Get-ProfileLastUse -Sid $p.SID -LocalPath $p.LocalPath
        if (-not $lastUse) { Write-Log "  Skipped $name (last use unknown)" 'WARN'; continue }
        if ($lastUse -ge $cutoff) { continue }

        $desc = '{0} (last used {1:yyyy-MM-dd})' -f $p.LocalPath, $lastUse
        if ($PSCmdlet.ShouldProcess($desc, 'Remove user profile')) {
            try {
                Remove-CimInstance -InputObject $p -ErrorAction Stop
                Write-Log "  Removed: $desc"
                $removed++
            }
            catch { Write-Log "  Failed to remove ${desc}: $($_.Exception.Message)" 'ERROR' }
        }
    }
    Write-Log "STEP 3 done: $removed profile(s) removed"
}

#endregion

#region Main ---------------------------------------------------------------------------------

try {
    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force -WhatIf:$false | Out-Null
    }
    $script:LogFile = Join-Path $LogDirectory ('Cleanup_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))

    # Keep only the 30 most recent logs
    Get-ChildItem -LiteralPath $LogDirectory -Filter 'Cleanup_*.log' |
        Sort-Object LastWriteTime -Descending | Select-Object -Skip 30 |
        Remove-Item -Force -WhatIf:$false -ErrorAction SilentlyContinue

    Write-Log ("=== Disk cleanup started on {0}{1} ===" -f $env:COMPUTERNAME, $(if ($WhatIfPreference) { ' (WhatIf)' } else { '' }))
    $startFree = Get-FreeSpaceGB

    if (-not (Test-BelowThreshold) -and -not $Force) {
        Write-Log 'Enough free space, nothing to do.'
        exit 0
    }

    $steps = @(
        @{ Skip = $SkipTemp;          Run = { Invoke-TempCleanup } },
        @{ Skip = $SkipWindowsUpdate; Run = { Invoke-WindowsUpdateCleanup } },
        @{ Skip = $SkipProfiles;      Run = { Invoke-OldProfileCleanup } }
    )

    foreach ($step in $steps) {
        if ($step.Skip) { continue }
        try { & $step.Run }
        catch { Write-Log "Step failed: $($_.Exception.Message)" 'ERROR' }

        if (-not $Force -and -not $WhatIfPreference -and -not (Test-BelowThreshold)) {
            Write-Log 'Threshold reached, remaining steps skipped.'
            break
        }
    }

    $endFree = Get-FreeSpaceGB
    Write-Log ("=== Done. Free: {0} GB -> {1} GB (+{2:N2} GB). Log: {3} ===" -f `
        $startFree, $endFree, ($endFree - $startFree), $script:LogFile)

    if ($endFree -lt $MinFreeGB) {
        Write-Log 'C: is still below the threshold after cleanup.' 'WARN'
        exit 1
    }
    exit 0
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    exit 2
}

#endregion
