<#
.SYNOPSIS
    One line on what the check looks for and what the repair does.

.DESCRIPTION
    With -DetectOnly the script only checks: it exits with 1 when the device needs
    fixing, so Intune runs the remediation, and with 0 when it doesn't. Without
    -DetectOnly it repairs. The line written to the output shows up in the device
    status in Intune.

    Intune runs remediation scripts without arguments, so upload a copy with
    -DetectOnly set to $true as the detection script.

.PARAMETER DetectOnly
    Only check, change nothing.

.NOTES
    Author:  Jonathan van Zijverden
    Version: 1.0.0
    Context: SYSTEM, 64-bit PowerShell
#>
[CmdletBinding()]
param (
    [switch]$DetectOnly
)

$ErrorActionPreference = 'Stop'

$ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$LogFolder = Join-Path -Path $env:ProgramData -ChildPath $ScriptName
$LogFile = Join-Path -Path $LogFolder -ChildPath "$ScriptName.log"

function Write-Log {
    param (
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    if (-not (Test-Path -Path $LogFolder)) {
        New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
    }
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message" | Out-File -FilePath $LogFile -Append -Encoding utf8
}

function Test-NeedsFix {
    # Replace with the actual check
    return $false
}

try {
    if (-not (Test-NeedsFix)) {
        Write-Output 'OK'
        exit 0
    }

    if ($DetectOnly) {
        Write-Log -Message 'Needs fixing' -Level 'WARN'
        Write-Output 'Not compliant: <what is wrong>'
        exit 1
    }

    Write-Log -Message 'Repairing'
    # Replace with the actual repair

    Write-Log -Message 'Repaired'
    Write-Output 'Fixed'
    exit 0
}
catch {
    Write-Log -Message "Failed: $($_.Exception.Message)" -Level 'ERROR'
    Write-Output "Failed: $($_.Exception.Message)"
    exit 1
}
