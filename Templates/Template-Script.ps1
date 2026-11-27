<#
.SYNOPSIS
    One line on what the script does.

.DESCRIPTION
    What the script changes on the device and when to use it.

.NOTES
    Author:  Jonathan van Zijverden
    Version: 1.0.0
    Context: SYSTEM, 64-bit PowerShell
#>
[CmdletBinding()]
param ()

$ErrorActionPreference = 'Stop'

$ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$LogFolder = Join-Path -Path $env:ProgramData -ChildPath $ScriptName
$LogFile = Join-Path -Path $LogFolder -ChildPath "$ScriptName.log"

function Write-Log {
    param (
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'WARNING', 'ERROR')]
        [string]$Level = 'INFO'
    )

    if (-not (Test-Path -Path $LogFolder)) {
        New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
    }
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message" | Out-File -FilePath $LogFile -Append -Encoding utf8
}

try {
    Write-Log -Message 'Started'

    # Work goes here

    Write-Log -Message 'Finished'
    exit 0
}
catch {
    Write-Log -Message "Failed: $($_.Exception.Message)" -Level 'ERROR'
    exit 1
}
