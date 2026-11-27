<#
.SYNOPSIS
    Custom compliance discovery script: is Sophos installed and running?

.DESCRIPTION
    Finds the Sophos services by display name ('Sophos*') and returns:
      SophosInstalled : at least one Sophos service exists
      SophosRunning   : every Sophos service set to start automatically is running

.NOTES
    Version: 1.0.0
    Context: SYSTEM, 64-bit PowerShell (Intune custom compliance)
#>

$ServicePattern = 'Sophos*'

$services = @(Get-Service -DisplayName $ServicePattern -ErrorAction SilentlyContinue)
$automatic = @($services | Where-Object { $_.StartType -eq 'Automatic' })
$stopped = @($automatic | Where-Object { $_.Status -ne 'Running' })

$result = @{
    SophosInstalled = ($services.Count -gt 0)
    SophosRunning   = ($automatic.Count -gt 0 -and $stopped.Count -eq 0)
}

return $result | ConvertTo-Json -Compress
