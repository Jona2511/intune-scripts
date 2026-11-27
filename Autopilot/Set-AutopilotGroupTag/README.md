# Set-AutopilotGroupTag

Bulk-changes Autopilot group tags from a CSV of serial numbers.

## Description

```text
Reads a CSV with the column SerialNumber and optionally GroupTag. With -GroupTag the
same tag is applied to every serial number (overrides a GroupTag column).
The input delimiter (';', ',' or tab) is detected automatically.

For each serial number the script looks up the windowsAutopilotDeviceIdentity:
  - up to -LookupThreshold serials: one filtered query per serial
    (contains(serialNumber,'...'), then an exact match);
  - more serials: all Autopilot devices are fetched once and indexed by serial.

The tag is changed with POST .../windowsAutopilotDeviceIdentities/{id}/updateDeviceProperties.
Supports -WhatIf.

Result per serial (CSV): Updated, Unchanged, NotFound, Duplicate, Failed, Skipped,
or WouldUpdate (under -WhatIf).
  Duplicate = serial found more than once in Autopilot, or listed in the CSV more
              than once with different tags. Such serials are never changed.
  Skipped   = empty group tag in the CSV.
```

## Parameters

- `-TenantId`: Tenant ID or tenant domain (e.g. contoso.onmicrosoft.com).
- `-CsvPath`: CSV with the column SerialNumber and optionally GroupTag.
- `-GroupTag`: One group tag for every serial in the CSV. Overrides the GroupTag column.
- `-ClientId`: App registration (client) ID for app-only sign-in with a certificate. Without -ClientId the script signs in interactively (delegated).
- `-CertificateThumbprint`: Thumbprint of the certificate (CurrentUser\My or LocalMachine\My) for -ClientId.
- `-OutputDirectory`: Folder for the result CSV and the log. Default: current folder.
- `-Delimiter`: Delimiter of the result CSV. Default ';' (Dutch Excel).
- `-LookupThreshold`: Up to this number of serials a filtered query per serial is used; above it, all Autopilot devices are fetched once.

## Examples

```powershell
.\Set-AutopilotGroupTag.ps1 -TenantId contoso.onmicrosoft.com -CsvPath .\tags.csv -WhatIf
```

```powershell
.\Set-AutopilotGroupTag.ps1 -TenantId $tid -ClientId $app -CertificateThumbprint $thumb `
    -CsvPath .\serials.csv -GroupTag 'NL-Office'
```

## Notes

```text
Graph permissions
  Application : DeviceManagementServiceConfig.ReadWrite.All
                Organization.Read.All (optional, for the tenant domain in file names)
  Delegated   : DeviceManagementServiceConfig.ReadWrite.All, User.Read
                (role: Intune Administrator or equivalent)

Dependency: Microsoft.Graph.Authentication.
The new tag shows in Intune within minutes. The [OrderID] on the Entra device, and
therefore dynamic group membership, follows later.

Exit codes: 0 = all Updated/Unchanged, 1 = NotFound/Duplicate/Failed/Skipped present,
2 = fatal error.
```
