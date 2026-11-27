#Requires -Version 5.1

<#
.SYNOPSIS
    Universal winget template: install, upgrade, uninstall or detect any app as SYSTEM.

.DESCRIPTION
    Built for Intune Win32 apps (and RMM) that run as SYSTEM, where 'winget' is not on PATH.

      - Resolves winget.exe from %ProgramFiles%\WindowsApps\Microsoft.DesktopAppInstaller_*
        (newest version, matching CPU architecture). Falls back to 'winget' on PATH when
        running as a user.
      - Detection: 'winget list --id <AppId> --exact'. The line containing the exact Id is
        parsed by tokens (no column headers, so locale independent). If winget truncated
        the Id (ellipsis), 'winget export --include-versions' is used instead. Optional extra
        fallback: uninstall registry keys via -DisplayNamePattern.
      - Install: 'winget install' (upgrades an existing older install). If the package has
        no installer for the requested scope, it retries without --scope.
      - Intune detection mode (-DetectOnly): writes ONE line to stdout and exits 0 only when
        the app is installed with a version >= the required version; otherwise writes
        nothing to stdout (log file only).

    Intune Win32 app command lines (package this script with IntuneWinAppUtil):
      Install  : powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-WingetApp.ps1 -AppId 7zip.7zip
      Uninstall: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-WingetApp.ps1 -AppId 7zip.7zip -Uninstall
      Detection: custom script = a copy of this script with the default values of
                 -AppId / -MinimumVersion filled in and -DetectOnly set to $true
                 (Intune detection scripts cannot take parameters).
      Return codes: 0 = success, 3010 = soft reboot, 1 = failed.

.PARAMETER AppId
    winget package identifier, e.g. '7zip.7zip' or 'Microsoft.PowerShell'.

.PARAMETER Version
    Install exactly this version (pin). Detection then requires at least this version.

.PARAMETER MinimumVersion
    Detection/compliance requires at least this version. Without -Version or
    -MinimumVersion any installed version counts as detected.

.PARAMETER Scope
    'machine' (default), 'user' or 'none' (no --scope argument).

.PARAMETER Override
    Replaces all installer arguments (winget --override).

.PARAMETER Custom
    Appends arguments to the default silent arguments (winget --custom).

.PARAMETER Source
    winget source (default 'winget').

.PARAMETER DisplayNamePattern
    Optional wildcard for the uninstall registry DisplayName, used as detection fallback
    when winget cannot be run or cannot find the package.

.PARAMETER Uninstall
    Uninstalls the package.

.PARAMETER DetectOnly
    Intune detection mode (see description).

.EXAMPLE
    .\Invoke-WingetApp.ps1 -AppId 'Notepad++.Notepad++' -MinimumVersion 8.6

.EXAMPLE
    .\Invoke-WingetApp.ps1 -AppId 'Zoom.Zoom' -Custom 'ZoomAutoUpdate=1'

.NOTES
    Exit codes (install/uninstall): 0 = success, 3010 = reboot required, 1 = failed.
    Exit codes (-DetectOnly)      : 0 + stdout = detected, 1 + no stdout = not detected.
    Log: %ProgramData%\WingetApp\<AppId>_*.log  (winget's own log: <AppId>_winget_*.log)
    Under SYSTEM, winget needs the Visual C++ 2015-2022 x64 runtime; without it winget.exe
    fails with 0xC0000135.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [ValidatePattern('^[\w\.\-\+]+$')]
    [string]$AppId,

    [string]$Version,

    [string]$MinimumVersion,

    [ValidateSet('machine', 'user', 'none')]
    [string]$Scope = 'machine',

    [string]$Override,

    [string]$Custom,

    [string]$Source = 'winget',

    [string]$DisplayNamePattern,

    [switch]$Uninstall,

    [switch]$DetectOnly,

    [string]$LogDirectory = (Join-Path $env:ProgramData 'WingetApp')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# winget HRESULTs (APPINSTALLER_CLI_ERROR_*), as signed Int32 hex strings
$WG = @{
    NoApplicableInstaller = '0x8A150010'
    NoPackagesFound       = '0x8A150014'
    UpdateNotApplicable   = '0x8A15002B'
    PackageAlreadyInstalled = '0x8A150061'
    PackageInUse          = '0x8A150101'
    RebootToFinish        = '0x8A150109'
    RebootToInstall       = '0x8A15010A'
    DllNotFound           = '0xC0000135'
}

#region Helpers ------------------------------------------------------------------------------

$script:LogFile = $null
$script:Quiet   = [bool]$DetectOnly   # Intune detection: nothing on stdout unless detected

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    if (-not $script:Quiet) {
        switch ($Level) {
            'WARN'  { Write-Host $line -ForegroundColor Yellow }
            'ERROR' { Write-Host $line -ForegroundColor Red }
            default { Write-Host $line }
        }
    }
    if ($script:LogFile) {
        try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -WhatIf:$false } catch { }
    }
}

function ConvertTo-Hex([int]$Code) { '0x{0:X8}' -f $Code }

function Compare-AppVersion {
    <#
        Compares two version strings part by part (numeric where possible).
        Returns -1, 0 or 1. Handles '1.2.3-beta', '23.01', '158.0.4370.45'.
    #>
    param([string]$A, [string]$B)
    $pa = @($A -split '[\.\-\+_ ]' | Where-Object { $_ -ne '' })
    $pb = @($B -split '[\.\-\+_ ]' | Where-Object { $_ -ne '' })
    $n = [math]::Max($pa.Count, $pb.Count)
    for ($i = 0; $i -lt $n; $i++) {
        $x = if ($i -lt $pa.Count) { $pa[$i] } else { '0' }
        $y = if ($i -lt $pb.Count) { $pb[$i] } else { '0' }
        $nx = 0L; $ny = 0L
        if ([long]::TryParse($x, [ref]$nx) -and [long]::TryParse($y, [ref]$ny)) {
            if ($nx -ne $ny) { return [math]::Sign($nx - $ny) }
        }
        else {
            $c = [string]::Compare($x, $y, $true)
            if ($c -ne 0) { return [math]::Sign($c) }
        }
    }
    return 0
}

function Get-WingetPath {
    $arch = switch ($env:PROCESSOR_ARCHITECTURE) { 'ARM64' { 'arm64' } 'x86' { 'x86' } default { 'x64' } }
    $root = Join-Path $env:ProgramFiles 'WindowsApps'
    $dirs = @(Get-ChildItem -LiteralPath $root -Directory -Filter "Microsoft.DesktopAppInstaller_*_${arch}__8wekyb3d8bbwe" -ErrorAction SilentlyContinue)
    $best = $dirs | Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'winget.exe') } |
        Sort-Object { [version](($_.Name -split '_')[1]) } -Descending | Select-Object -First 1
    if ($best) { return (Join-Path $best.FullName 'winget.exe') }

    $cmd = Get-Command -Name winget.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

function Invoke-Winget {
    param([Parameter(Mandatory)][string[]]$Arguments)
    Write-Log ("  winget {0}" -f ($Arguments -join ' '))
    $prev = [Console]::OutputEncoding
    try {
        [Console]::OutputEncoding = [Text.Encoding]::UTF8
        $raw = & $script:Winget @Arguments 2>&1 | ForEach-Object { "$_" }
    }
    finally { [Console]::OutputEncoding = $prev }
    $code = $LASTEXITCODE
    # Strip progress spinners/bars
    # (progress bars consist of runs of block characters; match by code range, kept ASCII-only)
    $lines = @($raw | Where-Object { $_ -and $_ -notmatch '^\s*[-\\|/]\s*$' -and $_ -notmatch '[^\x00-\x7F]{3,}' })
    [pscustomobject]@{ ExitCode = $code; Hex = (ConvertTo-Hex $code); Lines = $lines }
}

function Get-InstalledVersionFromWinget {
    <# Returns the installed version string, or $null if not installed / unknown. #>
    $r = Invoke-Winget -Arguments @('list', '--id', $AppId, '--exact', '--source', $Source,
        '--accept-source-agreements', '--disable-interactivity')
    if ($r.Hex -eq $WG.NoPackagesFound) { return $null }
    if ($r.ExitCode -ne 0) {
        Write-Log "  winget list failed ($($r.Hex))" 'WARN'
        if ($r.Hex -eq $WG.DllNotFound) { Write-Log '  Visual C++ 2015-2022 runtime missing for winget under SYSTEM.' 'ERROR' }
        return $null
    }

    # Line with the exact Id as a whitespace-separated token; version is the next token
    $escaped = [regex]::Escape($AppId)
    foreach ($line in $r.Lines) {
        if ($line -notmatch "(^|\s)$escaped(\s|$)") { continue }
        $tokens = @($line -split '\s+' | Where-Object { $_ })
        $idx = [array]::IndexOf([string[]]$tokens, ($tokens | Where-Object { $_ -ieq $AppId } | Select-Object -First 1))
        if ($idx -ge 0 -and $idx + 1 -lt $tokens.Count) {
            $v = $tokens[$idx + 1]
            if ($v -in '<', '>' -and $idx + 2 -lt $tokens.Count) { $v = $tokens[$idx + 2] }
            return $v
        }
    }

    # Id probably truncated (ellipsis): use export
    Write-Log '  Id not found in list output (truncated?); using winget export.'
    $tmp = Join-Path $env:TEMP ("wgexport_{0}.json" -f [guid]::NewGuid())
    try {
        $e = Invoke-Winget -Arguments @('export', '-o', $tmp, '--include-versions', '--source', $Source,
            '--accept-source-agreements', '--disable-interactivity')
        if (Test-Path -LiteralPath $tmp) {
            $json = Get-Content -LiteralPath $tmp -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($src in @($json.Sources)) {
                foreach ($p in @($src.Packages)) {
                    if ($p.PackageIdentifier -ieq $AppId) { return [string]$p.Version }
                }
            }
        }
        elseif ($e.ExitCode -ne 0) { Write-Log "  winget export failed ($($e.Hex))" 'WARN' }
    }
    finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    return $null
}

function Get-InstalledVersionFromRegistry {
    if (-not $DisplayNamePattern) { return $null }
    $keys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    $hit = Get-ItemProperty -Path $keys -ErrorAction SilentlyContinue |
        Where-Object { $_.PSObject.Properties['DisplayName'] -and $_.DisplayName -like $DisplayNamePattern } |
        Select-Object -First 1
    if ($hit -and $hit.PSObject.Properties['DisplayVersion']) { return [string]$hit.DisplayVersion }
    if ($hit) { return '0' }
    return $null
}

function Get-InstalledVersion {
    $v = $null
    if ($script:Winget) { $v = Get-InstalledVersionFromWinget }
    if (-not $v) {
        $rv = Get-InstalledVersionFromRegistry
        if ($rv) { Write-Log "  Found via registry ($DisplayNamePattern): $rv"; $v = $rv }
    }
    return $v
}

function Test-VersionOk([string]$Installed) {
    if (-not $Installed) { return $false }
    $required = if ($Version) { $Version } elseif ($MinimumVersion) { $MinimumVersion } else { $null }
    if (-not $required) { return $true }
    if ($Installed -in 'Unknown', 'Onbekend') { return $false }
    return ((Compare-AppVersion $Installed $required) -ge 0)
}

#endregion

#region Main ---------------------------------------------------------------------------------

try {
    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force -WhatIf:$false | Out-Null
    }
    $safeId = $AppId -replace '[^\w\.\-]', '_'
    $script:LogFile = Join-Path $LogDirectory ('{0}_{1:yyyyMMdd_HHmmss}.log' -f $safeId, (Get-Date))
    Get-ChildItem -LiteralPath $LogDirectory -Filter "${safeId}_*.log" |
        Sort-Object LastWriteTime -Descending | Select-Object -Skip 30 |
        Remove-Item -Force -WhatIf:$false -ErrorAction SilentlyContinue

    $action = if ($DetectOnly) { 'Detect' } elseif ($Uninstall) { 'Uninstall' } else { 'Install' }
    $who = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    Write-Log "=== winget $action $AppId on $env:COMPUTERNAME as $who ==="

    $script:Winget = Get-WingetPath
    if ($script:Winget) { Write-Log "winget: $($script:Winget)" }
    else { Write-Log 'winget.exe not found (App Installer missing?).' 'ERROR' }

    # Detection -------------------------------------------------------------------------
    $installed = Get-InstalledVersion
    Write-Log ("Installed version: {0}" -f $(if ($installed) { $installed } else { '(not installed)' }))

    if ($DetectOnly) {
        if (Test-VersionOk $installed) {
            Write-Output "Detected: $AppId $installed"
            exit 0
        }
        exit 1
    }

    if (-not $script:Winget) { exit 1 }

    # Uninstall -------------------------------------------------------------------------
    if ($Uninstall) {
        if (-not $installed) { Write-Log 'Not installed; nothing to do.'; Write-Output "NOT INSTALLED: $AppId"; exit 0 }
        if (-not $PSCmdlet.ShouldProcess($AppId, 'winget uninstall')) { exit 0 }

        $r = Invoke-Winget -Arguments @('uninstall', '--id', $AppId, '--exact', '--source', $Source, '--silent',
            '--accept-source-agreements', '--disable-interactivity',
            '--log', (Join-Path $LogDirectory ('{0}_winget_{1:yyyyMMdd_HHmmss}.log' -f $safeId, (Get-Date))))
        switch ($r.Hex) {
            '0x00000000'          { Write-Log 'Uninstalled.'; Write-Output "UNINSTALLED: $AppId"; exit 0 }
            $WG.NoPackagesFound   { Write-Log 'Not found; treated as uninstalled.'; Write-Output "NOT INSTALLED: $AppId"; exit 0 }
            $WG.RebootToFinish    { Write-Log 'Uninstalled; reboot required.' 'WARN'; Write-Output "UNINSTALLED (reboot): $AppId"; exit 3010 }
            default {
                Write-Log ("Uninstall failed ({0}): {1}" -f $r.Hex, ($r.Lines | Select-Object -Last 3) -join ' | ') 'ERROR'
                Write-Output "FAILED: uninstall $AppId ($($r.Hex))"
                exit 1
            }
        }
    }

    # Install ---------------------------------------------------------------------------
    if (Test-VersionOk $installed) {
        Write-Log 'Already installed with a sufficient version.'
        Write-Output "COMPLIANT: $AppId $installed"
        exit 0
    }
    if (-not $PSCmdlet.ShouldProcess($AppId, 'winget install')) { exit 0 }

    $wgLog = Join-Path $LogDirectory ('{0}_winget_{1:yyyyMMdd_HHmmss}.log' -f $safeId, (Get-Date))
    $base = @('install', '--id', $AppId, '--exact', '--source', $Source, '--silent',
        '--accept-package-agreements', '--accept-source-agreements', '--disable-interactivity', '--log', $wgLog)
    if ($Version)  { $base += @('--version', $Version) }
    if ($Override) { $base += @('--override', $Override) }
    elseif ($Custom) { $base += @('--custom', $Custom) }

    $args1 = if ($Scope -ne 'none') { $base + @('--scope', $Scope) } else { $base }
    $r = Invoke-Winget -Arguments $args1
    if ($r.Hex -eq $WG.NoApplicableInstaller -and $Scope -ne 'none') {
        Write-Log "No installer for scope '$Scope'; retrying without --scope." 'WARN'
        $r = Invoke-Winget -Arguments $base
    }

    $result = switch ($r.Hex) {
        '0x00000000'                { 'OK' }
        $WG.UpdateNotApplicable     { 'OK' }      # already newest available
        $WG.PackageAlreadyInstalled { 'OK' }
        $WG.RebootToFinish          { 'REBOOT' }
        $WG.RebootToInstall         { 'REBOOT' }
        default                     { 'FAIL' }
    }

    if ($result -eq 'FAIL') {
        $hint = switch ($r.Hex) {
            $WG.NoPackagesFound       { 'package Id not found in source' }
            $WG.PackageInUse          { 'application in use' }
            $WG.DllNotFound           { 'Visual C++ runtime missing for winget' }
            default                   { ($r.Lines | Select-Object -Last 3) -join ' | ' }
        }
        Write-Log "Install failed ($($r.Hex)): $hint" 'ERROR'
        Write-Output "FAILED: $AppId ($($r.Hex)) $hint"
        exit 1
    }

    $after = Get-InstalledVersion
    Write-Log "Installed version after install: $after"
    if (-not (Test-VersionOk $after)) {
        Write-Log 'winget reported success but the required version is not detected.' 'WARN'
    }

    if ($result -eq 'REBOOT') {
        Write-Output "INSTALLED (reboot required): $AppId $after"
        exit 3010
    }
    Write-Output "INSTALLED: $AppId $after"
    exit 0
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    if (-not $DetectOnly) { Write-Output "ERROR: $($_.Exception.Message)" }
    exit 1
}

#endregion
