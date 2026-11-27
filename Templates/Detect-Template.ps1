<#
.SYNOPSIS
    Detection script for an Intune remediation.

.DESCRIPTION
    Exits with 1 when the device needs fixing, so Intune runs the remediation script,
    and with 0 when the device is fine. The line written to the output shows up in
    the device status in Intune.

.NOTES
    Author:  Jonathan van Zijverden
    Version: 1.0.0
    Context: SYSTEM, 64-bit PowerShell
#>
[CmdletBinding()]
param ()

try {
    # Replace with the actual check
    $NeedsFix = $false

    if ($NeedsFix) {
        Write-Output 'Not compliant: <what is wrong>'
        exit 1
    }

    Write-Output 'Compliant'
    exit 0
}
catch {
    Write-Output "Detection failed: $($_.Exception.Message)"
    exit 1
}
