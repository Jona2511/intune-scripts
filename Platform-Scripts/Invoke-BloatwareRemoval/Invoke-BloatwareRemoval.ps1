#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Removes provisioned Windows apps (bloatware) that are not on an allowlist.

.DESCRIPTION
    - Reads all provisioned Appx packages (Get-AppxProvisionedPackage -Online). These are
      installed for every new user profile.
    - Every package whose DisplayName does not match the allowlist is removed with
      Remove-AppxProvisionedPackage, so new users no longer get it.
    - Framework packages (VCLibs, UI.Xaml, .NET Native, WindowsAppRuntime) and packages
      Windows marks as NonRemovable are always skipped.
    - With -IncludeInstalled the same apps are also removed for existing users
      (Get-AppxPackage -AllUsers | Remove-AppxPackage -AllUsers). Only Microsoft-published
      packages (and the provisioned apps found above) are considered there, so apps users
      installed themselves from other publishers are left alone.

    Default allowlist (wildcards on the package name): Store, Store purchase, App Installer
    (winget), Terminal, Calculator, Photos, Notepad, Paint, Snipping Tool, Camera, Windows
    Security, codec extensions (HEIF, HEVC, VP9, WebP, AV1, Raw, MPEG2, Web Media), Company
    Portal, new Teams, Quick Assist, Edge, winget source and framework/runtime packages.
    Extend it with -AdditionalAllowList or -AllowListPath, or replace it with -AllowList.

.PARAMETER AllowList
    Replaces the default allowlist.

.PARAMETER AdditionalAllowList
    Added to the (default) allowlist, e.g. 'Microsoft.ZuneMusic','Microsoft.OutlookForWindows'.

.PARAMETER AllowListPath
    Text file with extra allowed names, one per line ('#' = comment).

.PARAMETER IncludeInstalled
    Also removes the apps for existing user profiles.

.PARAMETER DetectOnly
    Only report. Exit code 1 if apps outside the allowlist are present.

.EXAMPLE
    .\Invoke-BloatwareRemoval.ps1 -DetectOnly

.EXAMPLE
    .\Invoke-BloatwareRemoval.ps1 -AdditionalAllowList 'Microsoft.ZuneMusic' -IncludeInstalled -WhatIf

.NOTES
    Exit codes: 0 = nothing (left) to remove, 1 = apps remain / removal failed, 2 = fatal error.
    Log: %ProgramData%\BloatwareRemoval\
    Removed apps can be reinstalled from the Microsoft Store or winget.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string[]]$AllowList = @(
        'Microsoft.WindowsStore',
        'Microsoft.StorePurchaseApp',
        'Microsoft.DesktopAppInstaller',
        'Microsoft.WindowsTerminal',
        'Microsoft.WindowsCalculator',
        'Microsoft.Windows.Photos',
        'Microsoft.WindowsNotepad',
        'Microsoft.Paint',
        'Microsoft.ScreenSketch',
        'Microsoft.WindowsCamera',
        'Microsoft.SecHealthUI',
        'Microsoft.HEIFImageExtension',
        'Microsoft.HEVCVideoExtension',
        'Microsoft.VP9VideoExtensions',
        'Microsoft.WebpImageExtension',
        'Microsoft.AV1VideoExtension',
        'Microsoft.RawImageExtension',
        'Microsoft.MPEG2VideoExtension',
        'Microsoft.WebMediaExtensions',
        'Microsoft.AVCEncoderVideoExtension',
        'Microsoft.CompanyPortal',
        'MSTeams',
        'MicrosoftCorporationII.QuickAssist',
        'Microsoft.MicrosoftEdge*',
        'Microsoft.Winget.Source',
        'Microsoft.ApplicationCompatibilityEnhancements',
        'Microsoft.StartExperiencesApp',
        'Microsoft.WidgetsPlatformRuntime',
        'Microsoft.VCLibs*',
        'Microsoft.UI.Xaml*',
        'Microsoft.NET.Native*',
        'Microsoft.WindowsAppRuntime*',
        'Microsoft.DirectXRuntime'
    ),

    [string[]]$AdditionalAllowList = @(),

    [string]$AllowListPath,

    [switch]$IncludeInstalled,
    [switch]$DetectOnly,

    [string]$LogDirectory = (Join-Path $env:ProgramData 'BloatwareRemoval')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Never touched, regardless of the allowlist
$FrameworkPattern = '^Microsoft\.(VCLibs|UI\.Xaml|NET\.Native|WindowsAppRuntime|DirectXRuntime|Services\.Store\.Engagement)'

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

function Test-Allowed([string]$Name, [string[]]$List) {
    foreach ($a in $List) { if ($Name -like $a) { return $true } }
    return $false
}

#endregion

#region Main ---------------------------------------------------------------------------------

try {
    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force -WhatIf:$false | Out-Null
    }
    $script:LogFile = Join-Path $LogDirectory ('Bloatware_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
    Get-ChildItem -LiteralPath $LogDirectory -Filter 'Bloatware_*.log' |
        Sort-Object LastWriteTime -Descending | Select-Object -Skip 30 |
        Remove-Item -Force -WhatIf:$false -ErrorAction SilentlyContinue

    $mode = if ($DetectOnly) { 'DetectOnly' } elseif ($WhatIfPreference) { 'WhatIf' } else { 'Remediate' }
    Write-Log "=== Bloatware removal on $env:COMPUTERNAME ($mode) ==="

    # 1. Allowlist -----------------------------------------------------------------------
    $allow = New-Object System.Collections.Generic.List[string]
    foreach ($a in @($AllowList) + @($AdditionalAllowList)) { if ($a -and $a.Trim()) { $allow.Add($a.Trim()) } }
    if ($AllowListPath) {
        if (-not (Test-Path -LiteralPath $AllowListPath)) { throw "Allowlist not found: $AllowListPath" }
        Get-Content -LiteralPath $AllowListPath -Encoding UTF8 |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -and -not $_.StartsWith('#') } |
            ForEach-Object { $allow.Add($_) }
    }
    Write-Log ("Allowlist: {0} entries" -f $allow.Count)

    # Names that Windows marks as framework / non-removable for any user
    $installed = @(Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue)
    $protected = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($pkg in $installed) {
        if ($pkg.IsFramework -or $pkg.NonRemovable -or "$($pkg.SignatureKind)" -eq 'System') { [void]$protected.Add($pkg.Name) }
    }

    # 2. Provisioned packages ------------------------------------------------------------
    $provisioned = @(Get-AppxProvisionedPackage -Online)
    $toRemove = @($provisioned | Where-Object {
        -not (Test-Allowed $_.DisplayName $allow) -and
        $_.DisplayName -notmatch $FrameworkPattern -and
        -not $protected.Contains($_.DisplayName)
    } | Sort-Object DisplayName)

    Write-Log ("Provisioned: {0} package(s), {1} not on allowlist" -f $provisioned.Count, $toRemove.Count)
    foreach ($p in $toRemove) { Write-Log "  Provisioned bloat: $($p.DisplayName) $($p.Version)" 'WARN' }

    # 3. Installed packages for existing users ------------------------------------------
    $installedToRemove = @()
    if ($IncludeInstalled) {
        $bloatNames = @($toRemove | ForEach-Object { $_.DisplayName })
        $installedToRemove = @($installed | Where-Object {
            -not $_.IsFramework -and -not $_.NonRemovable -and "$($_.SignatureKind)" -ne 'System' -and
            $_.Name -notmatch $FrameworkPattern -and
            -not (Test-Allowed $_.Name $allow) -and
            ( $bloatNames -contains $_.Name -or
              $_.PackageFamilyName -like '*_8wekyb3d8bbwe' -or
              $_.Name -like 'Clipchamp.*' -or
              $_.Name -like 'MicrosoftCorporationII.*' -or
              $_.Name -like 'MicrosoftWindows.*' )
        } | Sort-Object Name, Version -Unique)
        foreach ($p in $installedToRemove) { Write-Log "  Installed bloat: $($p.Name) $($p.Version)" 'WARN' }
    }

    $total = $toRemove.Count + $installedToRemove.Count
    if ($total -eq 0) {
        Write-Output 'COMPLIANT: no apps outside the allowlist'
        exit 0
    }
    $summary = (@($toRemove | ForEach-Object DisplayName) + @($installedToRemove | ForEach-Object Name) |
        Sort-Object -Unique) -join ', '
    if ($DetectOnly) {
        Write-Output "NON-COMPLIANT: $total package(s) outside allowlist: $summary"
        exit 1
    }

    # 4. Remove ----------------------------------------------------------------------------
    $failed = New-Object System.Collections.Generic.List[string]
    $hasAllUsers = (Get-Command Remove-AppxProvisionedPackage).Parameters.ContainsKey('AllUsers')

    foreach ($p in $toRemove) {
        if (-not $PSCmdlet.ShouldProcess($p.DisplayName, 'Remove provisioned package')) { $failed.Add($p.DisplayName); continue }
        try {
            if ($hasAllUsers) { Remove-AppxProvisionedPackage -Online -PackageName $p.PackageName -AllUsers | Out-Null }
            else              { Remove-AppxProvisionedPackage -Online -PackageName $p.PackageName | Out-Null }
            Write-Log "  Removed provisioned: $($p.DisplayName)"
        }
        catch {
            Write-Log "  Failed provisioned $($p.DisplayName): $($_.Exception.Message)" 'ERROR'
            $failed.Add($p.DisplayName)
        }
    }

    foreach ($p in $installedToRemove) {
        if (-not $PSCmdlet.ShouldProcess($p.PackageFullName, 'Remove installed package for all users')) { $failed.Add($p.Name); continue }
        try {
            Remove-AppxPackage -Package $p.PackageFullName -AllUsers -ErrorAction Stop
            Write-Log "  Removed installed: $($p.Name) $($p.Version)"
        }
        catch {
            Write-Log "  Failed installed $($p.Name): $($_.Exception.Message)" 'ERROR'
            $failed.Add($p.Name)
        }
    }

    if ($failed.Count -gt 0) {
        $msg = if ($WhatIfPreference) { 'NON-COMPLIANT (WhatIf, nothing changed)' } else { 'NON-COMPLIANT: not removed' }
        Write-Output ("{0}: {1}" -f $msg, (($failed | Sort-Object -Unique) -join ', '))
        exit 1
    }
    Write-Output "REMEDIATED: removed $summary"
    exit 0
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 2
}

#endregion
