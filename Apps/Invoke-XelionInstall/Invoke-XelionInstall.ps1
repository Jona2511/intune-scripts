#Requires -Version 5.1

<#
.SYNOPSIS
    Silent install, uninstall and Intune detection for the Xelion softphone (Windows app).

.DESCRIPTION
    Xelion distributes its Windows app through the Microsoft Store and apps.xelion.com.
    Verified on a reference device: the current Windows app is an MSIX/Store package
    named '71F23AFD.Xelion' (publisher id c91qzwq3fg5hw, Store Id 9NBLGGH5Q9BJ).
    The separate "Xelion Admin" tool is a per-user ClickOnce app and is excluded from
    detection by default (-ExcludeDisplayNamePattern).

    The installer format can differ per release, so this script supports all three
    common formats and picks the method from the file extension:

      - MSIX / MSIXBUNDLE / APPX / APPXBUNDLE:
          Provisioned for all users with Add-AppxProvisionedPackage -SkipLicense.
          Existing users get the app at their next sign-in; new users at first sign-in.
      - MSI:
          msiexec /i <file> /qn /norestart /l*v <log>  (+ -MsiProperties)
      - EXE:
          <file> -SilentArgs (no universal silent switch exists for EXE installers:
          check the vendor's switches and set -SilentArgs).

    Without -InstallerPath the script uses the first Xelion* installer it finds next to
    itself (the Intune .intunewin content folder).

    Detection (installed + version >= -MinimumVersion), all of these are checked:
      - provisioned and installed AppX/MSIX packages whose name matches -AppxNamePattern
      - uninstall registry entries (HKLM 64/32-bit and the loaded HKU user hives)
        whose DisplayName matches -DisplayNamePattern

    Store alternative (no installer file needed): Intune "Microsoft Store app (new)" with
    Id 9NBLGGH5Q9BJ, or Invoke-WingetApp.ps1 -AppId 9NBLGGH5Q9BJ -Source msstore.
    This script's -DetectOnly detects that Store install as well.

    Intune Win32 app command lines:
      Install  : powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-XelionInstall.ps1
      Uninstall: powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Invoke-XelionInstall.ps1 -Uninstall
      Detection: custom script = copy of this script with -DetectOnly / -MinimumVersion
                 defaults filled in (detection scripts cannot take parameters).

.PARAMETER InstallerPath
    Path to the Xelion installer (.msix/.msixbundle/.appx/.appxbundle/.msi/.exe).

.PARAMETER SilentArgs
    Arguments for an .exe installer (default '/S' - verify for your installer version).

.PARAMETER MsiProperties
    Extra MSI properties, e.g. 'ALLUSERS=1'.

.PARAMETER MinimumVersion
    Minimum version for detection/compliance. Empty = any version.

.PARAMETER CloseProcesses
    Process names (wildcards) to close before install/uninstall. Empty array = do not close.

.PARAMETER AppxNamePattern
    Wildcard for the AppX/MSIX package name.

.PARAMETER DisplayNamePattern
    Wildcard for the uninstall registry DisplayName.

.PARAMETER ExcludeDisplayNamePattern
    DisplayName wildcards to ignore (default: 'Xelion Admin*').

.PARAMETER Uninstall
    Removes Xelion (AppX/MSIX for all users + deprovision, MSI, or EXE uninstaller).

.PARAMETER UninstallArgs
    Arguments for an EXE uninstaller without QuietUninstallString (default '/S').

.PARAMETER DetectOnly
    Intune detection mode: stdout + exit 0 when detected, nothing on stdout otherwise.

.EXAMPLE
    .\Invoke-XelionInstall.ps1 -InstallerPath .\XelionSetup.msix -MinimumVersion 8.4

.EXAMPLE
    .\Invoke-XelionInstall.ps1 -DetectOnly -MinimumVersion 8.4

.NOTES
    Exit codes: 0 = success, 3010 = reboot required, 1 = failed.
    -DetectOnly: 0 + stdout = detected, 1 + no stdout = not detected.
    Run as SYSTEM/administrator. Log: %ProgramData%\XelionInstall\
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$InstallerPath,

    [string]$SilentArgs = '/S',

    [string]$MsiProperties,

    [string]$MinimumVersion,

    [string[]]$CloseProcesses = @('Xelion*'),

    [string]$AppxNamePattern = '*.Xelion',

    [string]$DisplayNamePattern = 'Xelion*',

    [string[]]$ExcludeDisplayNamePattern = @('Xelion Admin*'),

    [switch]$Uninstall,

    [string]$UninstallArgs = '/S',

    [switch]$DetectOnly,

    [string]$LogDirectory = (Join-Path $env:ProgramData 'XelionInstall')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Helpers ------------------------------------------------------------------------------

$script:LogFile = $null
$script:Quiet   = [bool]$DetectOnly

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

function Compare-AppVersion {
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

function Get-XelionInstalls {
    <# Returns all Xelion installs found: Kind (Appx/AppxProvisioned/Registry), Name, Version, details. #>
    $found = New-Object System.Collections.Generic.List[object]

    try {
        foreach ($p in @(Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like $AppxNamePattern })) {
            $found.Add([pscustomobject]@{ Kind = 'AppxProvisioned'; Name = $p.DisplayName; Version = [string]$p.Version; Id = $p.PackageName; Scope = 'All users' })
        }
    }
    catch { Write-Log "  Provisioned packages not readable: $($_.Exception.Message)" 'WARN' }

    try {
        # -AllUsers needs elevation; as a normal user fall back to the current user's packages
        $appx = try { @(Get-AppxPackage -AllUsers -Name $AppxNamePattern) } catch { @(Get-AppxPackage -Name $AppxNamePattern) }
        foreach ($p in $appx) {
            $users = @($p.PackageUserInformation | Where-Object { "$($_.InstallState)" -eq 'Installed' } |
                ForEach-Object { "$($_.UserSecurityId.Username)" })
            $found.Add([pscustomobject]@{ Kind = 'Appx'; Name = $p.Name; Version = [string]$p.Version; Id = $p.PackageFullName; Scope = ($users -join ', ') })
        }
    }
    catch { Write-Log "  AppX packages not readable: $($_.Exception.Message)" 'WARN' }

    $roots = @(
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall';             Scope = 'Machine (64-bit)' },
        @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'; Scope = 'Machine (32-bit)' }
    )
    foreach ($hive in @(Get-ChildItem -LiteralPath 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue |
            Where-Object { $_.PSChildName -match '^S-1-(5-21|12-1)-[\d\-]+$' })) {
        $roots += @{ Path = "Registry::HKEY_USERS\$($hive.PSChildName)\Software\Microsoft\Windows\CurrentVersion\Uninstall"; Scope = "User $($hive.PSChildName)" }
    }
    foreach ($r in $roots) {
        if (-not (Test-Path -LiteralPath $r.Path)) { continue }
        foreach ($k in @(Get-ChildItem -LiteralPath $r.Path -ErrorAction SilentlyContinue)) {
            $v = Get-ItemProperty -LiteralPath $k.PSPath -ErrorAction SilentlyContinue
            if (-not $v -or -not $v.PSObject.Properties['DisplayName'] -or $v.DisplayName -notlike $DisplayNamePattern) { continue }
            if ($ExcludeDisplayNamePattern | Where-Object { $v.DisplayName -like $_ }) { continue }
            $get = { param($n) if ($v.PSObject.Properties[$n]) { [string]$v.$n } else { $null } }
            $found.Add([pscustomobject]@{
                Kind            = 'Registry'
                Name            = $v.DisplayName
                Version         = (& $get 'DisplayVersion')
                Id              = $k.PSChildName
                Scope           = $r.Scope
                UninstallString = (& $get 'UninstallString')
                QuietUninstall  = (& $get 'QuietUninstallString')
                IsMsi           = ((& $get 'WindowsInstaller') -eq '1' -or $k.PSChildName -match '^\{[0-9A-Fa-f\-]{36}\}$')
            })
        }
    }
    return ,$found
}

function Test-Compliant($Installs) {
    foreach ($i in $Installs) {
        if (-not $MinimumVersion) { return $true }
        if ($i.Version -and (Compare-AppVersion $i.Version $MinimumVersion) -ge 0) { return $true }
    }
    return $false
}

function Stop-XelionProcesses {
    foreach ($pattern in $CloseProcesses) {
        foreach ($p in @(Get-Process -Name $pattern -ErrorAction SilentlyContinue)) {
            Write-Log "  Closing process $($p.ProcessName) (PID $($p.Id))"
            try { $p | Stop-Process -Force -ErrorAction Stop } catch { Write-Log "  Could not close $($p.ProcessName): $($_.Exception.Message)" 'WARN' }
        }
    }
}

function Resolve-Installer {
    if ($InstallerPath) {
        if (-not (Test-Path -LiteralPath $InstallerPath)) { throw "Installer not found: $InstallerPath" }
        return (Resolve-Path -LiteralPath $InstallerPath).ProviderPath
    }
    $dir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $candidate = Get-ChildItem -LiteralPath $dir -File |
        Where-Object { $_.Name -like 'Xelion*' -and $_.Extension -in '.msix', '.msixbundle', '.appx', '.appxbundle', '.msi', '.exe' } |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $candidate) { throw "No Xelion installer found next to the script ($dir). Use -InstallerPath." }
    return $candidate.FullName
}

function Install-Xelion([string]$Path) {
    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    $sig = Get-AuthenticodeSignature -LiteralPath $Path
    Write-Log ("Installer: {0} (signature: {1}, {2})" -f $Path, $sig.Status,
        $(if ($sig.SignerCertificate) { $sig.SignerCertificate.Subject } else { 'none' }))

    switch -Regex ($ext) {
        '^\.(msix|msixbundle|appx|appxbundle)$' {
            Write-Log 'Provisioning MSIX/AppX for all users...'
            Add-AppxProvisionedPackage -Online -PackagePath $Path -SkipLicense -ErrorAction Stop | Out-Null
            return 0
        }
        '^\.msi$' {
            $msiLog = Join-Path $LogDirectory ('msi_install_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
            $argList = "/i `"$Path`" /qn /norestart /l*v `"$msiLog`""
            if ($MsiProperties) { $argList += " $MsiProperties" }
            Write-Log "msiexec $argList"
            return (Start-Process -FilePath 'msiexec.exe' -ArgumentList $argList -Wait -PassThru).ExitCode
        }
        '^\.exe$' {
            Write-Log "Running $Path $SilentArgs"
            return (Start-Process -FilePath $Path -ArgumentList $SilentArgs -Wait -PassThru).ExitCode
        }
        default { throw "Unsupported installer type: $ext" }
    }
}

function Uninstall-Xelion {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param($Installs)
    $codes = @()
    foreach ($i in $Installs) {
        if (-not $PSCmdlet.ShouldProcess("$($i.Name) $($i.Version) [$($i.Kind), $($i.Scope)]", 'Uninstall')) { continue }
        try {
            switch ($i.Kind) {
                'AppxProvisioned' {
                    Remove-AppxProvisionedPackage -Online -PackageName $i.Id -ErrorAction Stop | Out-Null
                    Write-Log "  Deprovisioned $($i.Id)"
                }
                'Appx' {
                    Remove-AppxPackage -Package $i.Id -AllUsers -ErrorAction Stop
                    Write-Log "  Removed $($i.Id) for all users"
                }
                'Registry' {
                    if ($i.IsMsi) {
                        $msiLog = Join-Path $LogDirectory ('msi_uninstall_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
                        $code = (Start-Process msiexec.exe -ArgumentList "/x $($i.Id) /qn /norestart /l*v `"$msiLog`"" -Wait -PassThru).ExitCode
                    }
                    else {
                        $cmd = if ($i.QuietUninstall) { $i.QuietUninstall } else { $i.UninstallString }
                        if (-not $cmd) { throw 'No uninstall string' }
                        # Split "path with spaces.exe" args
                        if ($cmd -match '^\s*"([^"]+)"\s*(.*)$') { $exe = $Matches[1]; $a = $Matches[2] }
                        else { $exe, $a = $cmd -split '\s+', 2 }
                        if (-not $i.QuietUninstall) { $a = ("$a $UninstallArgs").Trim() }
                        Write-Log "  Running $exe $a"
                        $code = if ($a) { (Start-Process -FilePath $exe -ArgumentList $a -Wait -PassThru).ExitCode }
                                else    { (Start-Process -FilePath $exe -Wait -PassThru).ExitCode }
                    }
                    Write-Log "  Exit code: $code"
                    $codes += $code
                }
            }
        }
        catch {
            Write-Log "  Uninstall of $($i.Name) failed: $($_.Exception.Message)" 'ERROR'
            $codes += 1
        }
    }
    return ,$codes
}

#endregion

#region Main ---------------------------------------------------------------------------------

try {
    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force -WhatIf:$false | Out-Null
    }
    $script:LogFile = Join-Path $LogDirectory ('Xelion_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
    Get-ChildItem -LiteralPath $LogDirectory -Filter '*.log' |
        Sort-Object LastWriteTime -Descending | Select-Object -Skip 40 |
        Remove-Item -Force -WhatIf:$false -ErrorAction SilentlyContinue

    $action = if ($DetectOnly) { 'Detect' } elseif ($Uninstall) { 'Uninstall' } else { 'Install' }
    Write-Log "=== Xelion $action on $env:COMPUTERNAME ==="

    $installs = Get-XelionInstalls
    foreach ($i in $installs) { Write-Log ("Found: {0} {1} [{2}, {3}]" -f $i.Name, $i.Version, $i.Kind, $i.Scope) }
    if ($installs.Count -eq 0) { Write-Log 'Xelion not found.' }

    # Detection -------------------------------------------------------------------------
    if ($DetectOnly) {
        if ($installs.Count -gt 0 -and (Test-Compliant $installs)) {
            $best = $installs | Sort-Object { $_.Version } -Descending | Select-Object -First 1
            Write-Output "Detected: $($best.Name) $($best.Version) [$($best.Kind)]"
            exit 0
        }
        exit 1
    }

    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) { throw 'Install/uninstall requires administrator or SYSTEM.' }

    # Uninstall -------------------------------------------------------------------------
    if ($Uninstall) {
        if ($installs.Count -eq 0) { Write-Output 'NOT INSTALLED: Xelion'; exit 0 }
        Stop-XelionProcesses
        $codes = Uninstall-Xelion $installs
        if ($codes | Where-Object { $_ -notin 0, 1605, 3010, 1641 }) {
            Write-Output "FAILED: Xelion uninstall (exit codes $($codes -join ', '))"
            exit 1
        }
        if ($codes -contains 3010 -or $codes -contains 1641) { Write-Output 'UNINSTALLED (reboot): Xelion'; exit 3010 }
        Write-Output 'UNINSTALLED: Xelion'
        exit 0
    }

    # Install ---------------------------------------------------------------------------
    if ($installs.Count -gt 0 -and (Test-Compliant $installs)) {
        Write-Log 'Already installed with a sufficient version.'
        Write-Output 'COMPLIANT: Xelion already installed'
        exit 0
    }

    $installer = Resolve-Installer
    if (-not $PSCmdlet.ShouldProcess($installer, 'Install Xelion')) { exit 0 }

    Stop-XelionProcesses
    $code = Install-Xelion $installer
    Write-Log "Installer exit code: $code"

    if ($code -notin 0, 3010, 1641) {
        Write-Log 'Install failed.' 'ERROR'
        Write-Output "FAILED: Xelion install (exit code $code)"
        exit 1
    }

    $after = Get-XelionInstalls
    if (-not (Test-Compliant $after)) {
        Write-Log 'Installer finished but Xelion is not detected with the required version.' 'ERROR'
        Write-Output 'FAILED: Xelion not detected after install'
        exit 1
    }
    $best = $after | Sort-Object { $_.Version } -Descending | Select-Object -First 1
    if ($code -in 3010, 1641) {
        Write-Output "INSTALLED (reboot required): Xelion $($best.Version)"
        exit 3010
    }
    Write-Output "INSTALLED: Xelion $($best.Version) [$($best.Kind)]"
    exit 0
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    if (-not $DetectOnly) { Write-Output "ERROR: $($_.Exception.Message)" }
    exit 1
}

#endregion
