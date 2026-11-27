#Requires -Version 5.1

<#
.SYNOPSIS
    Cleans up the taskbar and Start menu on new Windows 11 devices by removing the
    standard Microsoft pins.

.DESCRIPTION
    What it configures (machine level, applies to NEW user profiles):

      - Taskbar: writes C:\Windows\OEM\TaskbarLayoutModification.xml with
        PinListPlacement="Replace" (only File Explorer plus -ExtraTaskbarPins) and registers
        it via HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\LayoutXMLPath.
        Windows applies it once, at a user's first sign-in. Users can change pins afterwards.

      - Start menu: copies a start2.bin (made on a reference machine with
        -ExportStartLayout) to the Default profile:
        C:\Users\Default\AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState
        New users get that pin layout at first sign-in.

      - Optional:
          -HideWidgets      policy  HKLM\SOFTWARE\Policies\Microsoft\Dsh\AllowNewsAndInterests = 0
          -SearchMode       policy  HKLM\SOFTWARE\Policies\Microsoft\Windows\Windows Search\SearchOnTaskbarMode
          -HideTaskView     Default user hive: Explorer\Advanced\ShowTaskViewButton = 0

      - Optional -ResetExistingUsers (best effort): for existing profiles, removes the
        taskbar pin state (Taskband key and User Pinned\TaskBar shortcuts) so the layout file
        is used again, and copies start2.bin into their profile. Explorer and the Start menu
        are restarted for users who are signed in.

    What is NOT supported / limitations (honest version):
      - Windows 11 has no API to unpin individual Start or taskbar items. This script
        replaces the whole layout instead.
      - start2.bin is an undocumented binary format. It is copied as-is and can break after
        a feature update; re-export it from a reference machine on the same build.
      - Several HKCU taskbar values (e.g. TaskbarDa for Widgets) are protected by the UCPD
        driver on recent builds and cannot be set by scripts. That is why Widgets and Search
        use policies here instead of HKCU values.
      - The taskbar layout is applied once per user. -ResetExistingUsers relies on the fact
        that Explorer rebuilds pins from the layout file when the pin state is missing; this
        works on current builds but is not documented behaviour.
      - Store-app pins (Store, Outlook, Copilot) live only in the Taskband state, so they
        disappear only through the layout file or a reset, not by deleting shortcuts.

.PARAMETER StartLayoutBin
    Path to a start2.bin exported with -ExportStartLayout. Without it, the Start menu is skipped.

.PARAMETER ExtraTaskbarPins
    Extra taskbar pins next to File Explorer. Per entry:
      - contains '!'           -> Store app AUMID, e.g. 'MSTeams_8wekyb3d8bbwe!MSTeams'
      - ends with '.lnk'       -> shortcut path, e.g. '%APPDATA%\Microsoft\Windows\Start Menu\Programs\Outlook.lnk'
      - otherwise              -> desktop application ID, e.g. 'MSEdge'

.PARAMETER ExportStartLayout
    Exports the current user's start2.bin to -ExportPath and stops. Run this as the user
    on a reference machine after arranging the Start menu manually.

.PARAMETER SearchMode
    Hide, Icon, IconAndLabel or Box. Not set = leave as is.

.EXAMPLE
    .\Invoke-StartTaskbarCleanup.ps1 -ExportStartLayout -ExportPath C:\Temp\start2.bin

.EXAMPLE
    .\Invoke-StartTaskbarCleanup.ps1 -StartLayoutBin .\start2.bin -HideWidgets -SearchMode Icon -HideTaskView

.EXAMPLE
    .\Invoke-StartTaskbarCleanup.ps1 -StartLayoutBin .\start2.bin -DetectOnly

.NOTES
    Exit codes: 0 = configured, 1 = not (fully) configured, 2 = fatal error.
    Log: %ProgramData%\StartTaskbarCleanup\
#>
[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Apply')]
param(
    [Parameter(ParameterSetName = 'Apply')]
    [string]$StartLayoutBin,

    [Parameter(ParameterSetName = 'Apply')]
    [string[]]$ExtraTaskbarPins = @(),

    [Parameter(ParameterSetName = 'Apply')]
    [switch]$HideWidgets,

    [Parameter(ParameterSetName = 'Apply')]
    [ValidateSet('Hide', 'Icon', 'IconAndLabel', 'Box')]
    [string]$SearchMode,

    [Parameter(ParameterSetName = 'Apply')]
    [switch]$HideTaskView,

    [Parameter(ParameterSetName = 'Apply')]
    [switch]$ResetExistingUsers,

    [Parameter(ParameterSetName = 'Apply')]
    [switch]$DetectOnly,

    [Parameter(ParameterSetName = 'Export', Mandatory)]
    [switch]$ExportStartLayout,

    [Parameter(ParameterSetName = 'Export')]
    [string]$ExportPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'start2.bin'),

    [string]$LogDirectory = (Join-Path $env:ProgramData 'StartTaskbarCleanup')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$StartPkgRel   = 'AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState'
$LayoutXmlPath = Join-Path $env:windir 'OEM\TaskbarLayoutModification.xml'
$ExplorerKey   = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer'
$DefaultProfile = Join-Path (Split-Path $env:PUBLIC -Parent) 'Default'
$SearchModes   = @{ Hide = 0; Icon = 1; IconAndLabel = 2; Box = 3 }

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
        try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 -WhatIf:$false } catch { }
    }
}

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function New-TaskbarLayoutXml {
    $pins = New-Object System.Collections.Generic.List[string]
    $pins.Add('        <taskbar:DesktopApp DesktopApplicationID="Microsoft.Windows.Explorer" />')
    foreach ($p in $ExtraTaskbarPins) {
        $e = [Security.SecurityElement]::Escape($p.Trim())
        if (-not $e) { continue }
        if ($p -match '!')          { $pins.Add("        <taskbar:UWA AppUserModelID=`"$e`" />") }
        elseif ($p -match '\.lnk$') { $pins.Add("        <taskbar:DesktopApp DesktopApplicationLinkPath=`"$e`" />") }
        else                        { $pins.Add("        <taskbar:DesktopApp DesktopApplicationID=`"$e`" />") }
    }
    @"
<?xml version="1.0" encoding="utf-8"?>
<LayoutModificationTemplate
    xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification"
    xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout"
    xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout"
    xmlns:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout"
    Version="1">
  <CustomTaskbarLayoutCollection PinListPlacement="Replace">
    <defaultlayout:TaskbarLayout>
      <taskbar:TaskbarPinList>
$($pins -join "`r`n")
      </taskbar:TaskbarPinList>
    </defaultlayout:TaskbarLayout>
  </CustomTaskbarLayoutCollection>
</LayoutModificationTemplate>
"@
}

function Get-RegValue([string]$Path, [string]$Name) {
    $p = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction SilentlyContinue
    if ($p) { return $p.$Name }
    return $null
}

function Set-RegDword([string]$Path, [string]$Name, [int]$Value) {
    if (-not (Test-Path -LiteralPath $Path)) { New-Item -Path $Path -Force | Out-Null }
    New-ItemProperty -LiteralPath $Path -Name $Name -Value $Value -PropertyType DWord -Force | Out-Null
}

function Invoke-WithHive {
    <# Loads an NTUSER.DAT under HKU\<TempName>, runs $Action with the registry root, unloads. #>
    param([string]$NtUserPath, [string]$TempName, [scriptblock]$Action)
    if (-not (Test-Path -LiteralPath $NtUserPath)) { Write-Log "  Hive not found: $NtUserPath" 'WARN'; return }
    $null = & reg.exe load "HKU\$TempName" $NtUserPath 2>&1
    if ($LASTEXITCODE -ne 0) { Write-Log "  Could not load $NtUserPath" 'WARN'; return }
    try { & $Action "Registry::HKEY_USERS\$TempName" }
    finally {
        [gc]::Collect(); [gc]::WaitForPendingFinalizers()
        $null = & reg.exe unload "HKU\$TempName" 2>&1
        if ($LASTEXITCODE -ne 0) { Write-Log "  Could not unload HKU\$TempName" 'WARN' }
    }
}

function Get-FileHashOrNull([string]$Path) {
    if (Test-Path -LiteralPath $Path) { return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash }
    return $null
}

#endregion

#region Export mode --------------------------------------------------------------------------

if ($ExportStartLayout) {
    $src = Join-Path $env:LOCALAPPDATA 'Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\start2.bin'
    if (-not (Test-Path -LiteralPath $src)) {
        Write-Output "ERROR: start2.bin not found for $env:USERNAME ($src)"
        exit 2
    }
    Copy-Item -LiteralPath $src -Destination $ExportPath -Force
    Write-Output "Exported start2.bin of $env:USERNAME to $ExportPath (build $([Environment]::OSVersion.Version))"
    exit 0
}

#endregion

#region Main ---------------------------------------------------------------------------------

try {
    if (-not (Test-IsAdmin)) { throw 'Run this script elevated or as SYSTEM (only -ExportStartLayout works as a normal user).' }

    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -Path $LogDirectory -ItemType Directory -Force -WhatIf:$false | Out-Null
    }
    $script:LogFile = Join-Path $LogDirectory ('StartTaskbar_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
    Get-ChildItem -LiteralPath $LogDirectory -Filter 'StartTaskbar_*.log' |
        Sort-Object LastWriteTime -Descending | Select-Object -Skip 30 |
        Remove-Item -Force -WhatIf:$false -ErrorAction SilentlyContinue

    $mode = if ($DetectOnly) { 'DetectOnly' } elseif ($WhatIfPreference) { 'WhatIf' } else { 'Remediate' }
    Write-Log "=== Start/taskbar cleanup on $env:COMPUTERNAME ($mode) ==="

    if ($StartLayoutBin -and -not (Test-Path -LiteralPath $StartLayoutBin)) {
        throw "start2.bin not found: $StartLayoutBin"
    }

    $desiredXml   = New-TaskbarLayoutXml
    $defaultStart = Join-Path (Join-Path $DefaultProfile $StartPkgRel) 'start2.bin'
    $searchPolKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search'
    $dshKey       = 'HKLM:\SOFTWARE\Policies\Microsoft\Dsh'

    # 1. Detection -------------------------------------------------------------------------
    function Get-Findings {
        $f = New-Object System.Collections.Generic.List[string]

        $currentXml = if (Test-Path -LiteralPath $LayoutXmlPath) { (Get-Content -LiteralPath $LayoutXmlPath -Raw).Trim() } else { $null }
        if ($currentXml -ne $desiredXml.Trim())                           { $f.Add('taskbar layout file missing/different') }
        if ((Get-RegValue $ExplorerKey 'LayoutXMLPath') -ne $LayoutXmlPath) { $f.Add('LayoutXMLPath not set') }

        if ($StartLayoutBin -and (Get-FileHashOrNull $defaultStart) -ne (Get-FileHashOrNull $StartLayoutBin)) {
            $f.Add('Default profile start2.bin missing/different')
        }
        if ($HideWidgets -and (Get-RegValue $dshKey 'AllowNewsAndInterests') -ne 0) { $f.Add('Widgets policy not set') }
        if ($SearchMode -and (Get-RegValue $searchPolKey 'SearchOnTaskbarMode') -ne $SearchModes[$SearchMode]) {
            $f.Add("search mode policy not $SearchMode")
        }
        if ($HideTaskView) {
            $script:tv = $null
            Invoke-WithHive -NtUserPath (Join-Path $DefaultProfile 'NTUSER.DAT') -TempName 'STC_Default' -Action {
                param($root)
                $script:tv = Get-RegValue (Join-Path $root 'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced') 'ShowTaskViewButton'
            }
            if ($script:tv -ne 0) { $f.Add('Task View not hidden in Default profile') }
        }
        return ,$f
    }

    $findings = Get-Findings
    foreach ($x in $findings) { Write-Log "  $x" 'WARN' }
    if (-not $StartLayoutBin) { Write-Log 'No -StartLayoutBin given: Start menu is not managed.' }

    if ($findings.Count -eq 0 -and -not $ResetExistingUsers) {
        Write-Output 'COMPLIANT: taskbar/Start defaults configured for new users'
        exit 0
    }
    if ($DetectOnly) {
        Write-Output ('NON-COMPLIANT: ' + ($findings -join '; '))
        exit 1
    }

    # 2. Apply ---------------------------------------------------------------------------
    if ($PSCmdlet.ShouldProcess($LayoutXmlPath, 'Write taskbar layout (Replace) and register LayoutXMLPath')) {
        New-Item -Path (Split-Path $LayoutXmlPath -Parent) -ItemType Directory -Force | Out-Null
        Set-Content -LiteralPath $LayoutXmlPath -Value $desiredXml -Encoding UTF8
        if (-not (Test-Path -LiteralPath $ExplorerKey)) { New-Item -Path $ExplorerKey -Force | Out-Null }
        New-ItemProperty -LiteralPath $ExplorerKey -Name 'LayoutXMLPath' -Value $LayoutXmlPath -PropertyType ExpandString -Force | Out-Null
        Write-Log "Taskbar layout written ($LayoutXmlPath)."
    }

    if ($StartLayoutBin -and $PSCmdlet.ShouldProcess($defaultStart, 'Copy start2.bin to Default profile')) {
        New-Item -Path (Split-Path $defaultStart -Parent) -ItemType Directory -Force | Out-Null
        Copy-Item -LiteralPath $StartLayoutBin -Destination $defaultStart -Force
        Write-Log 'start2.bin copied to Default profile.'
    }

    if ($HideWidgets -and $PSCmdlet.ShouldProcess('Widgets', 'Set policy AllowNewsAndInterests=0')) {
        Set-RegDword $dshKey 'AllowNewsAndInterests' 0
        Write-Log 'Widgets hidden (policy).'
    }

    if ($SearchMode -and $PSCmdlet.ShouldProcess('Taskbar search', "Set policy SearchOnTaskbarMode=$SearchMode")) {
        Set-RegDword $searchPolKey 'SearchOnTaskbarMode' $SearchModes[$SearchMode]
        Write-Log "Taskbar search mode set to $SearchMode (policy)."
    }

    if ($HideTaskView -and $PSCmdlet.ShouldProcess('Default profile', 'Hide Task View button')) {
        Invoke-WithHive -NtUserPath (Join-Path $DefaultProfile 'NTUSER.DAT') -TempName 'STC_Default' -Action {
            param($root)
            Set-RegDword (Join-Path $root 'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced') 'ShowTaskViewButton' 0
        }
        Write-Log 'Task View hidden for new users.'
    }

    # 3. Existing users (best effort) ----------------------------------------------------
    if ($ResetExistingUsers) {
        $profiles = Get-CimInstance -ClassName Win32_UserProfile -Filter 'Special=False' |
            Where-Object { $_.LocalPath -and $_.SID -match '^S-1-(5-21|12-1)-' }

        foreach ($p in $profiles) {
            $name = Split-Path $p.LocalPath -Leaf
            if (-not $PSCmdlet.ShouldProcess($name, 'Reset taskbar pins / Start layout (best effort)')) { continue }
            Write-Log "Resetting existing user $name"

            $resetAction = {
                param($root)
                $tb = Join-Path $root 'Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband'
                if (Test-Path -LiteralPath $tb) { Remove-Item -LiteralPath $tb -Recurse -Force }
                if ($HideTaskView) {
                    Set-RegDword (Join-Path $root 'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced') 'ShowTaskViewButton' 0
                }
            }
            try {
                if (Test-Path -LiteralPath "Registry::HKEY_USERS\$($p.SID)") { & $resetAction "Registry::HKEY_USERS\$($p.SID)" }
                else { Invoke-WithHive -NtUserPath (Join-Path $p.LocalPath 'NTUSER.DAT') -TempName "STC_$($p.SID)" -Action $resetAction }

                $pinned = Join-Path $p.LocalPath 'AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar'
                Get-ChildItem -LiteralPath $pinned -Filter '*.lnk' -ErrorAction SilentlyContinue | Remove-Item -Force

                if ($StartLayoutBin) {
                    $dst = Join-Path $p.LocalPath $StartPkgRel
                    if (Test-Path -LiteralPath $dst) { Copy-Item -LiteralPath $StartLayoutBin -Destination (Join-Path $dst 'start2.bin') -Force }
                }
            }
            catch { Write-Log "  Reset of $name failed: $($_.Exception.Message)" 'WARN' }
        }

        # Restart shell components so signed-in users see the change (Winlogon restarts Explorer)
        Get-Process -Name StartMenuExperienceHost, explorer -ErrorAction SilentlyContinue |
            Stop-Process -Force -ErrorAction SilentlyContinue
        Write-Log 'Explorer and Start menu restarted for signed-in users.'
    }

    if ($WhatIfPreference) {
        Write-Output ('NON-COMPLIANT (WhatIf, nothing changed): ' + ($findings -join '; '))
        exit 1
    }

    # 4. Verify ----------------------------------------------------------------------------
    $after = Get-Findings
    if ($after.Count -gt 0) {
        Write-Output ('NON-COMPLIANT after apply: ' + ($after -join '; '))
        exit 1
    }
    Write-Output 'REMEDIATED: taskbar/Start defaults configured for new users'
    exit 0
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 2
}

#endregion
