<#
.SYNOPSIS
    Custom compliance discovery script: counts installed apps that are not allowed.

.DESCRIPTION
    Searches the uninstall entries of the device (64-bit and 32-bit) and of every
    loaded user profile for display names that match $ForbiddenApps.
    Returns ForbiddenAppCount. Matches are also written to
    %ProgramData%\Test-ForbiddenApps\Test-ForbiddenApps.log, because a compliance
    script may only output the JSON.

.NOTES
    Version: 1.0.0
    Context: SYSTEM, 64-bit PowerShell (Intune custom compliance)
#>

$ForbiddenApps = @(
    'TeamViewer*'
    'AnyDesk*'
    '*uTorrent*'
    '*BitTorrent*'
)

$uninstallKeys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)
foreach ($sid in (Get-ChildItem -Path Registry::HKEY_USERS -ErrorAction SilentlyContinue).PSChildName) {
    if ($sid -match '^S-1-(5-21|12-1)-[\d-]+$') {
        $uninstallKeys += "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    }
}

$found = foreach ($key in $uninstallKeys) {
    foreach ($app in (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)) {
        foreach ($pattern in $ForbiddenApps) {
            if ($app.DisplayName -like $pattern) {
                $app.DisplayName
                break
            }
        }
    }
}
$found = @($found | Sort-Object -Unique)

$logFolder = Join-Path -Path $env:ProgramData -ChildPath 'Test-ForbiddenApps'
try {
    if (-not (Test-Path -Path $logFolder)) { New-Item -Path $logFolder -ItemType Directory -Force | Out-Null }
    $line = if ($found.Count) { "Found: $($found -join ', ')" } else { 'Nothing found' }
    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $line" | Out-File -FilePath (Join-Path $logFolder 'Test-ForbiddenApps.log') -Append -Encoding utf8
}
catch {
    Write-Verbose "Could not write the log: $($_.Exception.Message)"
}

return @{ ForbiddenAppCount = $found.Count } | ConvertTo-Json -Compress
