# Set-IntuneAssignments

Bulk-assigns Intune apps and policies to groups, from a CSV or parameters.

## Description

```text
Supported object types (-Type / CSV column Type):
  App               deviceAppManagement/mobileApps
  DeviceConfiguration  deviceManagement/deviceConfigurations   (templates, custom OMA-URI)
  Compliance        deviceManagement/deviceCompliancePolicies
  SettingsCatalog   deviceManagement/configurationPolicies     (incl. newer Endpoint security)
  AdminTemplate     deviceManagement/groupPolicyConfigurations
  PlatformScript    deviceManagement/deviceManagementScripts

The Intune /assign actions REPLACE all assignments of an object. This script therefore
reads the existing assignments first, merges the new ones in and posts the complete set,
so existing assignments are kept. Per object all rows are combined into one call.

Behaviour per row:
  - Group already assigned the same way          -> Unchanged
  - App already assigned to the group with another intent, or include <-> exclude
    switched                                       -> the existing assignment is replaced (Changed)
  - Otherwise                                      -> Added

Targets: group display name, group object ID, 'All Users' or 'All Devices'.
Optional assignment filter (by name) with FilterMode include/exclude.

CSV columns (delimiter -Delimiter, default ';'):
  Type;Name;Id;Group;Intent;Mode;Filter;FilterMode
    Name or Id        object display name or object ID
    Intent            apps only: required | available | uninstall (default required)
    Mode              include | exclude (default include; 'All Users/Devices' only include)
    Filter/FilterMode optional
```

## Parameters

- `-TenantId`: Tenant ID or primary domain.
- `-ClientId`: App registration for app-only auth (with -CertificateThumbprint). Omit for interactive login.
- `-CsvPath`: CSV with assignments. Alternative: -Type, -Name/-Id, -Group.
- `-Notifications`: For new Win32 app assignments: showAll | showReboot | hideAll (default showAll).

## Examples

```powershell
.\Set-IntuneAssignments.ps1 -TenantId contoso.onmicrosoft.com -CsvPath .\assignments.csv -WhatIf
```

```powershell
.\Set-IntuneAssignments.ps1 -TenantId contoso.onmicrosoft.com -Type App -Name '7-Zip','Google Chrome' `
    -Group 'SG-Intune-Devices-All' -Intent required -ClientId <appId> -CertificateThumbprint <thumb>
```

## Notes

```text
Required Graph permissions (application or delegated):
  DeviceManagementApps.ReadWrite.All
  DeviceManagementConfiguration.ReadWrite.All
  DeviceManagementScripts.ReadWrite.All   (platform scripts)
  Group.Read.All
Uses the beta endpoint for all Intune objects: settings catalog, admin templates,
platform scripts and assignment filters are not (fully) available in v1.0.
Remediations (deviceHealthScripts) are not supported: they need a run schedule per assignment.

Exit codes: 0 = all rows OK, 1 = one or more rows failed/not found, 2 = fatal error.
```
