<#
.SYNOPSIS
    Custom compliance discovery script: are all three Windows Firewall profiles on?

.DESCRIPTION
    Returns DomainEnabled, PrivateEnabled and PublicEnabled. A profile counts as on
    only when Windows reports it as enabled, so 'NotConfigured' counts as off.

.NOTES
    Version: 1.0.0
    Context: SYSTEM, 64-bit PowerShell (Intune custom compliance)
#>

$profiles = Get-NetFirewallProfile -PolicyStore ActiveStore

$result = @{}
foreach ($name in 'Domain', 'Private', 'Public') {
    $fwProfile = $profiles | Where-Object { $_.Name -eq $name }
    $result["${name}Enabled"] = ($null -ne $fwProfile -and $fwProfile.Enabled -eq 'True')
}

return $result | ConvertTo-Json -Compress
