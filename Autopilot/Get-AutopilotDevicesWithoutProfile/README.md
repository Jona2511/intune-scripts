# Get-AutopilotDevicesWithoutProfile

Reports Autopilot devices that have no (successfully) assigned deployment profile.

## Description

```text
Fetches all windowsAutopilotDeviceIdentities and reports every device whose
deploymentProfileAssignmentStatus is not "assigned...": notAssigned, pending, failed,
unknown. With -IncludeOutOfSync, assignedOutOfSync is reported as well.

For each reported device the script also reads the assigned and the intended profile
($expand=deploymentProfile,intendedDeploymentProfile), so you can see which profile is
on its way for a pending device. Skip this with -SkipProfileLookup on large sets.

Uses the BETA endpoint: deploymentProfileAssignmentStatus, the detailed status,
deploymentProfile and intendedDeploymentProfile do not exist in v1.0.
```

## Parameters

- `-TenantId`: Tenant ID or tenant domain.
- `-ClientId`: App registration (client) ID for app-only sign-in with a certificate. Without -ClientId the script signs in interactively (delegated).
- `-CertificateThumbprint`: Thumbprint of the certificate (CurrentUser\My or LocalMachine\My) for -ClientId.
- `-IncludeOutOfSync`: Also report devices with status assignedOutOfSync.
- `-SkipProfileLookup`: Skips the per-device lookup of the assigned/intended profile name.
- `-OutputDirectory`: Folder for the CSV and the log. Default: current folder.
- `-Delimiter`: CSV delimiter. Default ';' (Dutch Excel).

## Examples

```powershell
.\Get-AutopilotDevicesWithoutProfile.ps1 -TenantId contoso.onmicrosoft.com
```

```powershell
.\Get-AutopilotDevicesWithoutProfile.ps1 -TenantId $tid -ClientId $app -CertificateThumbprint $thumb -IncludeOutOfSync
```

## Notes

```text
Graph permissions
  Application : DeviceManagementServiceConfig.Read.All
                Organization.Read.All (optional, for the tenant domain in file names)
  Delegated   : DeviceManagementServiceConfig.Read.All, User.Read

Dependency: Microsoft.Graph.Authentication.
A freshly imported device usually stays 'pending' or 'notAssigned' for up to
15-30 minutes after its profile group membership is evaluated.

Exit codes: 0 = all devices have a profile, 1 = devices without a profile found,
2 = fatal error.
```
