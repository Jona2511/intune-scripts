#Requires -Version 5.1

<#
.SYNOPSIS
    Detects devices that have not rebooted for more than X days and shows the
    logged-on user a toast notification.

.DESCRIPTION
    Uptime is based on Win32_OperatingSystem.LastBootUpTime. With Fast Startup enabled
    a "Shut down" does not reset the uptime; only "Restart" does. The toast says so.

    The script works in two contexts:

      - As SYSTEM (Intune remediation, scheduled task, RMM):
        A toast cannot be shown from the SYSTEM session. The script copies itself to
        %ProgramData%\UptimeWarning, writes the toast texts to config.json and starts a
        temporary scheduled task that runs in the session of every logged-on user
        (BUILTIN\Users, interactive). That task runs this script with -ShowToast.
        The task is removed afterwards.

      - As the logged-on user:
        Shows the toast directly.

    The toast has two buttons:
      - "Nu herstarten": restarts after 60 seconds (cancel with 'shutdown /a').
        Uses a protocol handler (uptime-restart:) in HKCU of the user.
      - "Later": closes the notification.

.PARAMETER MaxUptimeDays
    Warn when the device has been up longer than this (default 14).

.PARAMETER Title
    Toast title.

.PARAMETER Message
    Toast text. {0} is replaced by the number of days.

.PARAMETER DetectOnly
    Only report, show no toast. Exit code 1 if the uptime is too long.

.PARAMETER ShowToast
    Internal: shows the toast in the current user session (used by the scheduled task).

.EXAMPLE
    .\Invoke-UptimeWarning.ps1 -DetectOnly

.EXAMPLE
    .\Invoke-UptimeWarning.ps1 -MaxUptimeDays 7

.NOTES
    Exit codes: 0 = uptime OK, 1 = uptime too long (toast shown or not), 2 = fatal error.
    Log (SYSTEM): %ProgramData%\UptimeWarning\
    Log (user)  : %LOCALAPPDATA%\UptimeWarning\
#>
[CmdletBinding()]
param(
    [ValidateRange(1, 365)]
    [int]$MaxUptimeDays = 14,

    [string]$Title = 'Tijd voor een herstart',

    [string]$Message = 'Je computer is al {0} dagen niet opnieuw opgestart. Herstart vandaag zodat updates worden afgerond en alles soepel blijft werken. Let op: kies "Opnieuw opstarten", want "Afsluiten" telt niet als herstart.',

    [switch]$DetectOnly,
    [switch]$ShowToast
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$AppFolder  = Join-Path $env:ProgramData 'UptimeWarning'
$ConfigFile = Join-Path $AppFolder 'config.json'
$TaskName   = 'UptimeWarning-Toast'
# AUMID of Windows PowerShell; registered on every Windows installation
$AppId      = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'

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
        try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 } catch { }
    }
}

function Initialize-Log([string]$Directory, [string]$Prefix) {
    if (-not (Test-Path -LiteralPath $Directory)) {
        New-Item -Path $Directory -ItemType Directory -Force | Out-Null
    }
    $script:LogFile = Join-Path $Directory ('{0}_{1:yyyyMMdd_HHmmss}.log' -f $Prefix, (Get-Date))
    Get-ChildItem -LiteralPath $Directory -Filter "${Prefix}_*.log" |
        Sort-Object LastWriteTime -Descending | Select-Object -Skip 30 |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

function Get-UptimeDays {
    $boot = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
    [pscustomobject]@{
        LastBoot = $boot
        Days     = [math]::Floor(((Get-Date) - $boot).TotalDays)
    }
}

function Test-IsSystem {
    [Security.Principal.WindowsIdentity]::GetCurrent().IsSystem
}

function Register-RestartProtocol {
    <# Registers uptime-restart: in HKCU so the toast button can trigger a restart. #>
    $base = 'HKCU:\Software\Classes\uptime-restart'
    $cmd  = 'shutdown.exe /r /t 60 /c "De computer wordt over 1 minuut opnieuw opgestart. Sla je werk op."'

    New-Item -Path "$base\shell\open\command" -Force | Out-Null
    Set-ItemProperty -LiteralPath $base -Name '(default)' -Value 'URL:Uptime restart'
    Set-ItemProperty -LiteralPath $base -Name 'URL Protocol' -Value ''
    Set-ItemProperty -LiteralPath "$base\shell\open\command" -Name '(default)' -Value $cmd
}

function Show-UptimeToast {
    param([string]$ToastTitle, [string]$ToastText)

    $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
    $null = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]

    $t = [Security.SecurityElement]::Escape($ToastTitle)
    $m = [Security.SecurityElement]::Escape($ToastText)

    # scenario="reminder" keeps the toast on screen until the user clicks a button
    $xml = @"
<toast scenario="reminder" launch="uptime-restart:" activationType="protocol">
  <visual>
    <binding template="ToastGeneric">
      <text>$t</text>
      <text>$m</text>
    </binding>
  </visual>
  <actions>
    <action content="Nu herstarten" activationType="protocol" arguments="uptime-restart:" />
    <action content="Later" activationType="system" arguments="dismiss" />
  </actions>
  <audio src="ms-winsoundevent:Notification.Reminder" />
</toast>
"@
    $doc = New-Object Windows.Data.Xml.Dom.XmlDocument
    $doc.LoadXml($xml)
    $toast = [Windows.UI.Notifications.ToastNotification]::new($doc)
    $toast.Tag   = 'UptimeWarning'
    $toast.Group = 'UptimeWarning'
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($AppId).Show($toast)
}

#endregion

#region User mode: show the toast ------------------------------------------------------------

if ($ShowToast) {
    try {
        Initialize-Log -Directory (Join-Path $env:LOCALAPPDATA 'UptimeWarning') -Prefix 'Toast'

        # Texts written by the SYSTEM part take precedence over the defaults
        if (Test-Path -LiteralPath $ConfigFile) {
            $cfg = Get-Content -LiteralPath $ConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json
            $Title   = $cfg.Title
            $Message = $cfg.Message
        }

        $uptime = Get-UptimeDays
        Register-RestartProtocol
        Show-UptimeToast -ToastTitle $Title -ToastText ($Message -f $uptime.Days)
        Write-Log "Toast shown to $env:USERNAME (uptime $($uptime.Days) days)."
        exit 0
    }
    catch {
        Write-Log "Showing toast failed: $($_.Exception.Message)" 'ERROR'
        exit 2
    }
}

#endregion

#region Main: detection ----------------------------------------------------------------------

try {
    $isSystem = Test-IsSystem
    if ($isSystem) { Initialize-Log -Directory $AppFolder -Prefix 'Uptime' }
    else           { Initialize-Log -Directory (Join-Path $env:LOCALAPPDATA 'UptimeWarning') -Prefix 'Uptime' }

    $uptime = Get-UptimeDays
    Write-Log ("=== Uptime check on {0}: last boot {1:yyyy-MM-dd HH:mm}, {2} days (limit {3}) ===" -f `
        $env:COMPUTERNAME, $uptime.LastBoot, $uptime.Days, $MaxUptimeDays)

    if ($uptime.Days -le $MaxUptimeDays) {
        Write-Output "COMPLIANT: uptime $($uptime.Days) days"
        exit 0
    }

    Write-Log "Uptime is longer than $MaxUptimeDays days." 'WARN'
    if ($DetectOnly) {
        Write-Output "NON-COMPLIANT: uptime $($uptime.Days) days (last boot $('{0:yyyy-MM-dd}' -f $uptime.LastBoot))"
        exit 1
    }

    # Running as a normal user: show the toast directly ----------------------------------
    if (-not $isSystem) {
        Register-RestartProtocol
        Show-UptimeToast -ToastTitle $Title -ToastText ($Message -f $uptime.Days)
        Write-Log "Toast shown to $env:USERNAME."
        Write-Output "NON-COMPLIANT: uptime $($uptime.Days) days - toast shown"
        exit 1
    }

    # Running as SYSTEM: hand over to the user session(s) via a scheduled task ---------
    $sessions = @(Get-CimInstance -ClassName Win32_Process -Filter "Name='explorer.exe'")
    if ($sessions.Count -eq 0) {
        Write-Log 'No user logged on; no toast shown.' 'WARN'
        Write-Output "NON-COMPLIANT: uptime $($uptime.Days) days - no user logged on"
        exit 1
    }

    if (-not (Test-Path -LiteralPath $AppFolder)) {
        New-Item -Path $AppFolder -ItemType Directory -Force | Out-Null
    }
    # Copy ourselves to a stable location (the Intune cache folder is temporary)
    $scriptCopy = Join-Path $AppFolder 'Invoke-UptimeWarning.ps1'
    if ($PSCommandPath -ne $scriptCopy) {
        Copy-Item -LiteralPath $PSCommandPath -Destination $scriptCopy -Force
    }
    [pscustomobject]@{ Title = $Title; Message = $Message } |
        ConvertTo-Json | Set-Content -LiteralPath $ConfigFile -Encoding UTF8

    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument (
        '-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "{0}" -ShowToast' -f $scriptCopy)
    # BUILTIN\Users, interactive: runs in the session of every logged-on user
    $principal = New-ScheduledTaskPrincipal -GroupId 'S-1-5-32-545' -RunLevel Limited
    $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -StartWhenAvailable

    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Register-ScheduledTask -TaskName $TaskName -Action $action -Principal $principal -Settings $settings `
        -Description 'Temporary task: shows the uptime warning to the logged-on user.' | Out-Null
    Start-ScheduledTask -TaskName $TaskName
    Write-Log "Scheduled task '$TaskName' started for $($sessions.Count) session(s)."

    # Wait until the task has finished, then remove it
    $deadline = (Get-Date).AddSeconds(60)
    do {
        Start-Sleep -Seconds 2
        $state = (Get-ScheduledTask -TaskName $TaskName).State
    } while ($state -eq 'Running' -and (Get-Date) -lt $deadline)

    $info = Get-ScheduledTaskInfo -TaskName $TaskName
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Log "Task finished (last result $($info.LastTaskResult)) and removed."

    Write-Output "NON-COMPLIANT: uptime $($uptime.Days) days - toast shown to logged-on user(s)"
    exit 1
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 2
}

#endregion
