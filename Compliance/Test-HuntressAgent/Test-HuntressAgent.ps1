<#
.SYNOPSIS
    Custom compliance discovery script: is Huntress installed and running?

.DESCRIPTION
    Finds the Huntress services by display name ('Huntress*') and returns:
      HuntressInstalled : at least one Huntress service exists
      HuntressRunning   : every Huntress service set to start automatically is running

.NOTES
    Version: 1.0.0
    Context: SYSTEM, 64-bit PowerShell (Intune custom compliance)
#>

$ServicePattern = 'Huntress*'

$services = @(Get-Service -DisplayName $ServicePattern -ErrorAction SilentlyContinue)
$automatic = @($services | Where-Object { $_.StartType -eq 'Automatic' })
$stopped = @($automatic | Where-Object { $_.Status -ne 'Running' })

$result = @{
    HuntressInstalled = ($services.Count -gt 0)
    HuntressRunning   = ($automatic.Count -gt 0 -and $stopped.Count -eq 0)
}

return $result | ConvertTo-Json -Compress
