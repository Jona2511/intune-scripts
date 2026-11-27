<#
.SYNOPSIS
    Remediation script for an Intune remediation.

.DESCRIPTION
    Runs when the detection script exits with 1. Exits with 0 when the fix worked
    and with 1 when it did not, so the result is visible in Intune.

.NOTES
    Author:  Jonathan van Zijverden
    Version: 1.0.0
    Context: SYSTEM, 64-bit PowerShell
#>
[CmdletBinding()]
param ()

try {
    # Replace with the actual fix

    Write-Output 'Fixed'
    exit 0
}
catch {
    Write-Output "Remediation failed: $($_.Exception.Message)"
    exit 1
}
