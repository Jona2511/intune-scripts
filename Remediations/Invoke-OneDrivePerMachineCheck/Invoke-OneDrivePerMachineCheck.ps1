#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Detects per-user OneDrive installs and moves the device to the per-machine client.

.DESCRIPTION
    Detection:
      - Per-machine install: %ProgramFiles%\Microsoft OneDrive\OneDrive.exe
        (or %ProgramFiles(x86)% on older installs).
      - Per-user installs: <profile>\AppData\Local\Microsoft\OneDrive\OneDrive.exe for
        every user profile.
      - Optional: per-machine version older than -MinimumVersion.

    Remediation (unless -DetectOnly):
      1. Per-machine client missing or too old -> runs OneDriveSetup.exe /allusers /silent.
         The installer comes from -InstallerPath (local/UNC) or is downloaded from -InstallerUrl.
      2. For every profile with a per-user install, the user's Run key (HKCU\...\Run\OneDrive)
         is pointed to the per-machine OneDrive.exe, so the user starts the per-machine
         client at the next sign-in. Hives of users who are not signed in are loaded
         temporarily.
      3. With -RemovePerUserBinaries: removes the per-user program files (OneDrive.exe,
         OneDriveStandaloneUpdater.exe and the version folders), the per-user uninstall
         entry and the per-user update task. Sync settings and data (the 'settings' folder
         and the OneDrive folder with files) are never touched. Skipped for users who have
         OneDrive running at that moment.

.PARAMETER InstallerPath
    Local or UNC path to OneDriveSetup.exe. Takes precedence over -InstallerUrl.

.PARAMETER InstallerUrl
    Download URL for OneDriveSetup.exe.

.PARAMETER MinimumVersion
    Optional: minimum per-machine version, e.g. '25.100.0.0'.

.PARAMETER RemovePerUserBinaries
    Also removes the per-user program files (see description).

.PARAMETER DetectOnly
    Only report. Exit code 1 if per-user installs exist or the per-machine client is missing.

.EXAMPLE
    .\Invoke-OneDrivePerMachineCheck.ps1 -DetectOnly

.EXAMPLE
    .\Invoke-OneDrivePerMachineCheck.ps1 -InstallerPath '\\fs01\sw$\OneDrive\OneDriveSetup.exe' -RemovePerUserBinaries

.NOTES
    Exit codes: 0 = per-machine only (afterwards), 1 = per-user installs remain / install failed,
    2 = fatal error.
    Log: %ProgramData%\OneDriveCheck\
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$InstallerPath,

    [string]$InstallerUrl = 'https://go.microsoft.com/fwlink/?linkid=844652',

    [version]$MinimumVersion,

    [switch]$RemovePerUserBinaries,
    [switch]$DetectOnly,

    [string]$LogDirectory = (Join-Path $env:ProgramData 'OneDriveCheck')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$RunKeyPath = 'Software\Microsoft\Windows\CurrentVersion\Run'

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

function Get-PerMachineOneDrive {
    $candidates = @(
        (Join-Path $env:ProgramFiles 'Microsoft OneDrive\OneDrive.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft OneDrive\OneDrive.exe')
    )
    foreach ($c in $candidates) {
        if ($c -and (Test-Path -LiteralPath $c)) {
            return [pscustomobject]@{
                Path    = $c
                Version = [version](Get-Item -LiteralPath $c).VersionInfo.FileVersion.Replace(',', '.').Trim()
            }
        }
    }
    return $null
}

function Get-PerUserInstalls {
    $profiles = Get-CimInstance -ClassName Win32_UserProfile -Filter 'Special=False' |
        Where-Object { $_.LocalPath -and $_.SID -match '^S-1-(5-21|12-1)-' }

    foreach ($p in $profiles) {
        $exe = Join-Path $p.LocalPath 'AppData\Local\Microsoft\OneDrive\OneDrive.exe'
        if (Test-Path -LiteralPath $exe) {
            $ver = $null
            try { $ver = [version](Get-Item -LiteralPath $exe).VersionInfo.FileVersion.Replace(',', '.').Trim() } catch { }
            [pscustomobject]@{
                Sid         = $p.SID
                ProfilePath = $p.LocalPath
                UserName    = Split-Path $p.LocalPath -Leaf
                Exe         = $exe
                Version     = $ver
                Loaded      = [bool]$p.Loaded
            }
        }
    }
}

function Invoke-WithUserHive {
    <#
        Runs $Action with the HKU root path of the user as argument. Loads NTUSER.DAT
        temporarily if the user is not signed in.
    #>
    param([Parameter(Mandatory)]$User, [Parameter(Mandatory)][scriptblock]$Action)

    if (Test-Path -LiteralPath "Registry::HKEY_USERS\$($User.Sid)") {
        & $Action "Registry::HKEY_USERS\$($User.Sid)"
        return
    }

    $tempKey = "OneDriveCheck_$($User.Sid)"
    $ntuser  = Join-Path $User.ProfilePath 'NTUSER.DAT'
    if (-not (Test-Path -LiteralPath $ntuser)) { Write-Log "  No NTUSER.DAT for $($User.UserName)" 'WARN'; return }

    $null = & reg.exe load "HKU\$tempKey" $ntuser 2>&1
    if ($LASTEXITCODE -ne 0) { Write-Log "  Could not load hive of $($User.UserName)" 'WARN'; return }
    try {
        & $Action "Registry::HKEY_USERS\$tempKey"
    }
    finally {
        [gc]::Collect(); [gc]::WaitForPendingFinalizers()
        $null = & reg.exe unload "HKU\$tempKey" 2>&1
        if ($LASTEXITCODE -ne 0) { Write-Log "  Could not unload hive of $($User.UserName)" 'WARN' }
    }
}

function Install-PerMachineOneDrive {
    $setup = $InstallerPath
    $downloaded = $false
    if (-not $setup) {
        $setup = Join-Path $env:TEMP 'OneDriveSetup.exe'
        Write-Log "Downloading OneDriveSetup.exe from $InstallerUrl"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $InstallerUrl -OutFile $setup -UseBasicParsing
        $downloaded = $true
    }
    if (-not (Test-Path -LiteralPath $setup)) { throw "Installer not found: $setup" }

    $sig = Get-AuthenticodeSignature -LiteralPath $setup
    if ($sig.Status -ne 'Valid' -or $sig.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
        throw "OneDriveSetup.exe signature is not a valid Microsoft signature ($($sig.Status))."
    }

    Write-Log "Running $setup /allusers /silent"
    $proc = Start-Process -FilePath $setup -ArgumentList '/allusers', '/silent' -Wait -PassThru
    if ($downloaded) { Remove-Item -LiteralPath $setup -Force -ErrorAction SilentlyContinue }
    Write-Log "  Installer exit code: $($proc.ExitCode)"
    return ($proc.ExitCode -eq 0)
}

#endregion

#region Main ---------------------------------------------------------------------------------

try {
    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force -WhatIf:$false | Out-Null
    }
    $script:LogFile = Join-Path $LogDirectory ('OneDrive_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
    Get-ChildItem -LiteralPath $LogDirectory -Filter 'OneDrive_*.log' |
        Sort-Object LastWriteTime -Descending | Select-Object -Skip 30 |
        Remove-Item -Force -WhatIf:$false -ErrorAction SilentlyContinue

    $mode = if ($DetectOnly) { 'DetectOnly' } elseif ($WhatIfPreference) { 'WhatIf' } else { 'Remediate' }
    Write-Log "=== OneDrive check on $env:COMPUTERNAME ($mode) ==="

    # 1. Detection -------------------------------------------------------------------------
    $machine = Get-PerMachineOneDrive
    if ($machine) { Write-Log "Per-machine OneDrive: $($machine.Version) ($($machine.Path))" }
    else          { Write-Log 'Per-machine OneDrive: not installed' 'WARN' }

    $machineOk = $machine -and (-not $MinimumVersion -or $machine.Version -ge $MinimumVersion)
    if ($machine -and -not $machineOk) {
        Write-Log "Per-machine version $($machine.Version) is older than $MinimumVersion." 'WARN'
    }

    $perUser = @(Get-PerUserInstalls)
    foreach ($u in $perUser) {
        Write-Log ("Per-user install: {0} version {1} ({2})" -f $u.UserName, $u.Version, $u.Exe) 'WARN'
    }

    if ($machineOk -and $perUser.Count -eq 0) {
        Write-Output "COMPLIANT: per-machine OneDrive $($machine.Version), no per-user installs"
        exit 0
    }

    $problems = @()
    if (-not $machine)   { $problems += 'per-machine client missing' }
    elseif (-not $machineOk) { $problems += "per-machine client $($machine.Version) < $MinimumVersion" }
    if ($perUser.Count)  { $problems += "per-user install(s): $(($perUser | ForEach-Object UserName) -join ', ')" }
    $problemText = $problems -join '; '

    if ($DetectOnly) {
        Write-Output "NON-COMPLIANT: $problemText"
        exit 1
    }

    # 2. Per-machine install -----------------------------------------------------------------
    if (-not $machineOk) {
        if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, 'Install OneDrive per-machine (/allusers)')) {
            $null = Install-PerMachineOneDrive
            $machine = Get-PerMachineOneDrive
            if (-not $machine) {
                Write-Log 'Per-machine install failed; per-user installs left as they are.' 'ERROR'
                Write-Output "NON-COMPLIANT: per-machine install failed ($problemText)"
                exit 1
            }
            Write-Log "Per-machine OneDrive installed: $($machine.Version)"
        }
        else {
            Write-Output "NON-COMPLIANT (WhatIf, nothing changed): $problemText"
            exit 1
        }
    }

    # 3. Move users to the per-machine client ---------------------------------------------
    $runValue = '"{0}" /background' -f $machine.Path
    $remaining = New-Object System.Collections.Generic.List[string]

    foreach ($u in $perUser) {
        Write-Log "User $($u.UserName):"

        if ($PSCmdlet.ShouldProcess($u.UserName, 'Point OneDrive Run key to per-machine client')) {
            Invoke-WithUserHive -User $u -Action {
                param($root)
                $run = Join-Path $root $RunKeyPath
                if (-not (Test-Path -LiteralPath $run)) { New-Item -Path $run -Force | Out-Null }
                Set-ItemProperty -LiteralPath $run -Name 'OneDrive' -Value $runValue
                Write-Log "  Run key -> $runValue"

                if ($RemovePerUserBinaries) {
                    $uninst = Join-Path $root 'Software\Microsoft\Windows\CurrentVersion\Uninstall\OneDriveSetup.exe'
                    if (Test-Path -LiteralPath $uninst) {
                        Remove-Item -LiteralPath $uninst -Recurse -Force
                        Write-Log '  Per-user uninstall entry removed'
                    }
                }
            }
        }

        if (-not $RemovePerUserBinaries) { $remaining.Add($u.UserName); continue }

        # Only when that user's OneDrive is not running
        $running = Get-CimInstance -ClassName Win32_Process -Filter "Name='OneDrive.exe'" |
            Where-Object { $_.ExecutablePath -and $_.ExecutablePath -like "$($u.ProfilePath)\*" }
        if ($running) {
            Write-Log '  Per-user OneDrive is running; binaries kept until the next run.' 'WARN'
            $remaining.Add($u.UserName)
            continue
        }

        $dir = Split-Path $u.Exe -Parent
        if ($PSCmdlet.ShouldProcess($dir, 'Remove per-user OneDrive program files')) {
            try {
                Get-ChildItem -LiteralPath $dir -Force | Where-Object {
                    $_.Name -in 'OneDrive.exe', 'OneDriveStandaloneUpdater.exe' -or
                    ($_.PSIsContainer -and $_.Name -match '^\d+\.\d+\.\d+\.\d+$')
                } | Remove-Item -Recurse -Force
                Write-Log "  Program files removed from $dir (settings kept)"

                Get-ScheduledTask -TaskName "OneDrive*$($u.Sid)*" -ErrorAction SilentlyContinue |
                    ForEach-Object {
                        Unregister-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -Confirm:$false
                        Write-Log "  Scheduled task removed: $($_.TaskName)"
                    }
            }
            catch {
                Write-Log "  Removal failed: $($_.Exception.Message)" 'ERROR'
                $remaining.Add($u.UserName)
            }
        }
        else { $remaining.Add($u.UserName) }
    }

    if ($remaining.Count -gt 0) {
        $msg = if ($RemovePerUserBinaries) { 'per-user binaries remain' } else { 'Run key redirected; per-user binaries kept (use -RemovePerUserBinaries)' }
        Write-Output ("NON-COMPLIANT: {0}: {1}" -f $msg, ($remaining -join ', '))
        exit 1
    }

    Write-Output "REMEDIATED: per-machine OneDrive $($machine.Version), per-user installs cleaned up"
    exit 0
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 2
}

#endregion
