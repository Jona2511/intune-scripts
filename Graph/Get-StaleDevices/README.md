# Get-StaleDevices

Reports Intune and Entra ID devices that have not checked in for more than X days.

## Description

```text
Report only: nothing is disabled or deleted.

  - Intune  : managedDevices with lastSyncDateTime older than -Days.
  - Entra ID: devices with approximateLastSignInDateTime older than -Days. A device
              that never signed in counts as stale when it was registered before
              the cutoff.

Both sources are matched on Intune azureADDeviceId = Entra deviceId. The combined
report shows one row per device with StaleIn = Intune, Entra or Both, and
Sources = Intune+Entra, IntuneOnly or EntraOnly.

Entra objects created by an Autopilot import (physicalIds contains [ZTDID]) are
marked IsAutopilotObject = True. Do not delete those as stale: they are needed for
future Autopilot deployments.

Output: three CSVs (Intune, Entra, combined) plus a log.
```

## Parameters

- `-TenantId`: Tenant ID or tenant domain.
- `-Days`: Number of days without check-in before a device is stale (default 90).
- `-ClientId`: App registration (client) ID for app-only sign-in with a certificate. Without -ClientId the script signs in interactively (delegated).
- `-CertificateThumbprint`: Thumbprint of the certificate (CurrentUser\My or LocalMachine\My) for -ClientId.
- `-OutputDirectory`: Folder for the CSVs and the log. Default: current folder.
- `-Delimiter`: CSV delimiter. Default ';' (Dutch Excel).

## Examples

```powershell
.\Get-StaleDevices.ps1 -TenantId contoso.onmicrosoft.com
```

```powershell
.\Get-StaleDevices.ps1 -TenantId $tid -ClientId $app -CertificateThumbprint $thumb -Days 120 -OutputDirectory D:\Reports
```

## Notes

```text
Graph permissions
  Application : DeviceManagementManagedDevices.Read.All, Device.Read.All
                Organization.Read.All (optional, for the tenant domain in file names)
  Delegated   : DeviceManagementManagedDevices.Read.All, Device.Read.All, User.Read

Dependency: Microsoft.Graph.Authentication.
approximateLastSignInDateTime is only updated about once every 14 days, so Entra
dates can lag by up to two weeks.
PrimaryUserUPN comes from managedDevice.userPrincipalName (the Intune primary user).

Exit codes: 0 = no stale devices, 1 = stale devices found, 2 = fatal error.
```
