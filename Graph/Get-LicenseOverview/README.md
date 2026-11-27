# Get-LicenseOverview

Licence overview for one tenant: subscriptions, assignments per user and users without a licence.

## Description

```text
Produces three CSV files in -OutputDirectory:

  Licenses_Skus_<domain>_<date>.csv         per SKU: enabled, consumed, unused, suspended,
                                            warning and locked-out units, with friendly names
  Licenses_Assignments_<domain>_<date>.csv  per user per SKU: direct and/or group-based
                                            (group names), assignment state and errors
  Licenses_Unlicensed_<domain>_<date>.csv   users without any licence (default: enabled
                                            member accounts only)

Friendly product names come from Microsoft's "Product names and service plan identifiers
for licensing" CSV (cached in %TEMP% for 7 days). If the download fails, the
skuPartNumber is used.

SKUs with 10,000 or more enabled units are free/viral offers (e.g. Power Automate Free,
Teams Exploratory). They are marked IsFreeOrViral and left out of the "unused" findings.
```

## Parameters

- `-TenantId`: Tenant ID or initial domain (contoso.onmicrosoft.com).
- `-ClientId`: App registration (client) ID for app-only authentication. Without it the script signs in interactively (delegated).
- `-CertificateThumbprint`: Thumbprint of the certificate (CurrentUser\My or LocalMachine\My) for app-only authentication.
- `-IncludeDisabled`: Also list disabled accounts in the unlicensed report.
- `-IncludeGuests`: Also list guest accounts in the unlicensed report.
- `-OutputDirectory`: Folder for CSVs and log. Default: current directory.
- `-Delimiter`: CSV delimiter. Default ';' (Dutch Excel).

## Examples

```powershell
.\Get-LicenseOverview.ps1 -TenantId contoso.onmicrosoft.com
```

```powershell
.\Get-LicenseOverview.ps1 -TenantId $tid -ClientId $appId -CertificateThumbprint $thumb -OutputDirectory C:\Reports
```

## Notes

```text
Required Graph permissions (application or delegated):
  Organization.Read.All   subscribedSkus and tenant name
  User.Read.All           users, assignedLicenses, licenseAssignmentStates
  Group.Read.All          names of licensing groups
(Directory.Read.All covers all three.)
Delegated: the signed-in user also needs a role such as Global Reader or License Administrator.

Dependency: Microsoft.Graph.Authentication.
Exit codes: 0 = no findings, 1 = findings (unused paid licences, unlicensed users,
licensed disabled accounts or assignment errors), 2 = fatal error.
```
