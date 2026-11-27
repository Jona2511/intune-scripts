#Requires -Version 5.1

<#
.SYNOPSIS
    Sets the time zone to W. Europe Standard Time, Dutch regional formats and home location,
    and keeps the keyboard on US International only.

.DESCRIPTION
    Target state:
      - Time zone            : W. Europe Standard Time (machine)
      - Regional format      : nl-NL (dates, numbers, currency) per user
      - Home location        : Netherlands (GeoId 176) per user
      - Keyboard             : ONLY US International (00020409) for every language in the
                               user's language list (e.g. 0409:00020409, 0413:00020409)
      - Display language     : not changed (the language list order stays as it is)
      - System locale        : nl-NL only with -SetSystemLocale (needs a reboot)

    The script works in two modes:

      - Machine mode (SYSTEM or elevated admin; default):
          1. Sets the time zone (and system locale with -SetSystemLocale).
          2. Applies the user settings to its own context and, if available (Windows 11),
             copies them with Copy-UserInternationalSettingsToSystem to the welcome screen
             and to new user profiles.
          3. Runs the per-user part for signed-in users through a temporary scheduled task
             for BUILTIN\Users (the script copies itself to %ProgramData%\RegionalSettings).
             Users who are not signed in get the settings at the next run, or via step 2 if
             their profile is new.

      - User mode (-UserMode, used by the scheduled task): applies format, home location and
        keyboard for the current user. Runs without elevation.

    Detection checks the time zone, the system locale (if requested), the welcome-screen
    defaults (HKU\.DEFAULT) and the settings of every signed-in user (via HKU).

.PARAMETER TimeZone
    Time zone ID (default 'W. Europe Standard Time').

.PARAMETER Culture
    Regional format (default 'nl-NL').

.PARAMETER GeoId
    Home location (default 176 = Netherlands).

.PARAMETER KeyboardLayout
    Keyboard layout ID kept for every language (default '00020409' = US International).

.PARAMETER SetSystemLocale
    Also sets the system locale (language for non-Unicode programs) to -Culture. Needs a reboot.

.PARAMETER DetectOnly
    Only report. Exit code 1 if something differs.

.PARAMETER UserMode
    Internal: applies the per-user settings in the current user session.

.EXAMPLE
    .\Invoke-RegionalSettings.ps1 -DetectOnly

.EXAMPLE
    .\Invoke-RegionalSettings.ps1 -SetSystemLocale

.NOTES
    Exit codes: 0 = in order (afterwards), 1 = differences remain, 2 = fatal error.
    Log (machine): %ProgramData%\RegionalSettings\
    Log (user)   : %LOCALAPPDATA%\RegionalSettings\
    Run in machine mode as an elevated admin, Copy-UserInternationalSettingsToSystem copies
    that admin's own settings (including the language list) to the welcome screen and new users.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$TimeZone = 'W. Europe Standard Time',

    [string]$Culture = 'nl-NL',

    [int]$GeoId = 176,

    [ValidatePattern('^[0-9A-Fa-f]{8}$')]
    [string]$KeyboardLayout = '00020409',

    [switch]$SetSystemLocale,
    [switch]$DetectOnly,
    [switch]$UserMode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$AppFolder = Join-Path $env:ProgramData 'RegionalSettings'
$TaskName  = 'RegionalSettings-User'
$KeyboardLayout = $KeyboardLayout.ToUpperInvariant()

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

function Initialize-Log([string]$Directory, [string]$Prefix) {
    if (-not (Test-Path -LiteralPath $Directory)) {
        New-Item -Path $Directory -ItemType Directory -Force -WhatIf:$false | Out-Null
    }
    $script:LogFile = Join-Path $Directory ('{0}_{1:yyyyMMdd_HHmmss}.log' -f $Prefix, (Get-Date))
    Get-ChildItem -LiteralPath $Directory -Filter "${Prefix}_*.log" |
        Sort-Object LastWriteTime -Descending | Select-Object -Skip 30 |
        Remove-Item -Force -WhatIf:$false -ErrorAction SilentlyContinue
}

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-UserHiveFindings {
    <#
        Checks one user hive (registry root such as Registry::HKEY_USERS\S-1-5-21-...).
        Returns a list of differences.
    #>
    param([string]$Root, [switch]$SkipKeyboard)
    $f = New-Object System.Collections.Generic.List[string]

    $intl = Get-ItemProperty -LiteralPath "$Root\Control Panel\International" -ErrorAction SilentlyContinue
    $locale = if ($intl -and $intl.PSObject.Properties['LocaleName']) { $intl.LocaleName } else { $null }
    if ($locale -ne $Culture) { $f.Add("format $locale") }

    $geo = Get-ItemProperty -LiteralPath "$Root\Control Panel\International\Geo" -ErrorAction SilentlyContinue
    $nation = if ($geo -and $geo.PSObject.Properties['Nation']) { "$($geo.Nation)" } else { $null }
    if ($nation -ne "$GeoId") { $f.Add("home location $nation") }

    if (-not $SkipKeyboard) {
        $profileKey = "$Root\Control Panel\International\User Profile"
        $tips = @()
        foreach ($lang in @(Get-ChildItem -LiteralPath $profileKey -ErrorAction SilentlyContinue)) {
            $tips += @($lang.GetValueNames() | Where-Object { $_ -match '^[0-9A-Fa-f]{4}:[0-9A-Fa-f{}-]+$' })
        }
        $wrong = @($tips | Where-Object { $_.Split(':')[1].ToUpperInvariant() -ne $KeyboardLayout })
        if ($tips.Count -eq 0)    { $f.Add('keyboard list unknown') }
        elseif ($wrong.Count -gt 0) { $f.Add("extra keyboard(s) $($wrong -join ',')") }
    }
    return ,$f
}

function Set-UserRegionalSettings {
    <# Applies format, home location and keyboard for the CURRENT user context. #>
    Set-Culture -CultureInfo $Culture
    Set-WinHomeLocation -GeoId $GeoId

    $list = Get-WinUserLanguageList
    foreach ($lang in $list) {
        # Use the language ID of the existing input method, otherwise derive it from the tag
        $langId = $null
        if ($lang.InputMethodTips.Count -gt 0) { $langId = $lang.InputMethodTips[0].Split(':')[0] }
        if (-not $langId) {
            try { $langId = '{0:X4}' -f ([Globalization.CultureInfo]::GetCultureInfo($lang.LanguageTag).LCID -band 0xFFFF) }
            catch { $langId = '0409' }
        }
        $lang.InputMethodTips.Clear()
        $lang.InputMethodTips.Add(('{0}:{1}' -f $langId, $KeyboardLayout))
    }
    Set-WinUserLanguageList -LanguageList $list -Force

    $firstTip = $list[0].InputMethodTips[0]
    Set-WinDefaultInputMethodOverride -InputTip $firstTip
    Write-Log ("User {0}: format {1}, home location {2}, keyboard {3} for {4} language(s)" -f `
        $env:USERNAME, $Culture, $GeoId, $KeyboardLayout, $list.Count)
}

#endregion

#region User mode ----------------------------------------------------------------------------

if ($UserMode) {
    try {
        Initialize-Log -Directory (Join-Path $env:LOCALAPPDATA 'RegionalSettings') -Prefix 'User'
        Set-UserRegionalSettings
        exit 0
    }
    catch {
        Write-Log "User mode failed: $($_.Exception.Message)" 'ERROR'
        exit 2
    }
}

#endregion

#region Machine mode -------------------------------------------------------------------------

try {
    if (-not (Test-IsAdmin)) { throw 'Machine mode needs SYSTEM or an elevated session.' }
    Initialize-Log -Directory $AppFolder -Prefix 'Regional'

    $isSystem = [Security.Principal.WindowsIdentity]::GetCurrent().IsSystem
    $mode = if ($DetectOnly) { 'DetectOnly' } elseif ($WhatIfPreference) { 'WhatIf' } else { 'Remediate' }
    Write-Log "=== Regional settings on $env:COMPUTERNAME ($mode, SYSTEM: $isSystem) ==="

    function Get-SignedInUserSids {
        Get-ChildItem -LiteralPath 'Registry::HKEY_USERS' |
            Where-Object { $_.PSChildName -match '^S-1-(5-21|12-1)-[\d-]+$' } |
            ForEach-Object { $_.PSChildName }
    }

    function Get-Findings {
        $f = New-Object System.Collections.Generic.List[string]
        $tz = (Get-TimeZone).Id
        if ($tz -ne $TimeZone) { $f.Add("time zone $tz") }
        if ($SetSystemLocale) {
            $sl = (Get-WinSystemLocale).Name
            if ($sl -ne $Culture) { $f.Add("system locale $sl") }
        }
        # Welcome screen / defaults; keyboard there is stored differently, so not checked
        foreach ($x in (Get-UserHiveFindings -Root 'Registry::HKEY_USERS\.DEFAULT' -SkipKeyboard)) { $f.Add("welcome screen: $x") }
        foreach ($sid in @(Get-SignedInUserSids)) {
            $name = $sid
            try { $name = (New-Object Security.Principal.SecurityIdentifier($sid)).Translate([Security.Principal.NTAccount]).Value } catch { }
            foreach ($x in (Get-UserHiveFindings -Root "Registry::HKEY_USERS\$sid")) { $f.Add("${name}: $x") }
        }
        return ,$f
    }

    # 1. Detection -------------------------------------------------------------------------
    $findings = Get-Findings
    foreach ($x in $findings) { Write-Log "  $x" 'WARN' }
    if ($findings.Count -eq 0) {
        Write-Output "COMPLIANT: $TimeZone, $Culture, GeoId $GeoId, keyboard $KeyboardLayout only"
        exit 0
    }
    if ($DetectOnly) {
        Write-Output ('NON-COMPLIANT: ' + ($findings -join '; '))
        exit 1
    }
    if (-not $PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Apply $TimeZone / $Culture / keyboard $KeyboardLayout")) {
        Write-Output ('NON-COMPLIANT (WhatIf, nothing changed): ' + ($findings -join '; '))
        exit 1
    }

    # 2. Machine settings ------------------------------------------------------------------
    if ((Get-TimeZone).Id -ne $TimeZone) {
        Set-TimeZone -Id $TimeZone
        Write-Log "Time zone set to $TimeZone."
    }
    if ($SetSystemLocale -and (Get-WinSystemLocale).Name -ne $Culture) {
        Set-WinSystemLocale -SystemLocale $Culture
        Write-Log "System locale set to $Culture (reboot required)." 'WARN'
    }

    # 3. Own context + welcome screen / new users ------------------------------------------
    Set-UserRegionalSettings
    if (Get-Command -Name Copy-UserInternationalSettingsToSystem -ErrorAction SilentlyContinue) {
        Copy-UserInternationalSettingsToSystem -WelcomeScreen $true -NewUser $true
        Write-Log 'Settings copied to welcome screen and new user profiles.'
    }
    else {
        Write-Log 'Copy-UserInternationalSettingsToSystem not available (needs Windows 11); welcome screen/new users not updated.' 'WARN'
    }

    # 4. Signed-in users via scheduled task -----------------------------------------------
    $sessions = @(Get-CimInstance -ClassName Win32_Process -Filter "Name='explorer.exe'")
    if ($sessions.Count -gt 0) {
        $scriptCopy = Join-Path $AppFolder 'Invoke-RegionalSettings.ps1'
        if ($PSCommandPath -ne $scriptCopy) { Copy-Item -LiteralPath $PSCommandPath -Destination $scriptCopy -Force }

        $taskArgs = '-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "{0}" -UserMode -Culture {1} -GeoId {2} -KeyboardLayout {3}' -f `
            $scriptCopy, $Culture, $GeoId, $KeyboardLayout
        $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $taskArgs
        $principal = New-ScheduledTaskPrincipal -GroupId 'S-1-5-32-545' -RunLevel Limited
        $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -StartWhenAvailable

        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
        Register-ScheduledTask -TaskName $TaskName -Action $action -Principal $principal -Settings $settings `
            -Description 'Temporary task: applies regional settings for signed-in users.' | Out-Null
        Start-ScheduledTask -TaskName $TaskName
        Write-Log "Per-user task started for $($sessions.Count) session(s)."

        $deadline = (Get-Date).AddSeconds(90)
        do {
            Start-Sleep -Seconds 3
            $state = (Get-ScheduledTask -TaskName $TaskName).State
        } while ($state -eq 'Running' -and (Get-Date) -lt $deadline)
        $info = Get-ScheduledTaskInfo -TaskName $TaskName
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Log "Per-user task finished (last result $($info.LastTaskResult)) and removed."
    }
    else {
        Write-Log 'No users signed in; per-user settings follow at their next run.'
    }

    # 5. Verify ----------------------------------------------------------------------------
    $after = Get-Findings
    if ($after.Count -gt 0) {
        foreach ($x in $after) { Write-Log "  Still different: $x" 'WARN' }
        Write-Output ('NON-COMPLIANT after apply: ' + ($after -join '; '))
        exit 1
    }
    Write-Output "REMEDIATED: $TimeZone, $Culture, GeoId $GeoId, keyboard $KeyboardLayout only"
    exit 0
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 2
}

#endregion
