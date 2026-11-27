# Test-NamedLocationTrusted

Checks whether one or more IP addresses (e.g. the office IP) fall inside a trusted
Conditional Access named location, and optionally fixes it.

## Description

```text
- Reads all IP named locations (/identity/conditionalAccess/namedLocations, v1.0).
- For every -IpAddress (IPv4/IPv6, single address or CIDR) it determines which named
  locations contain it (real CIDR math, not string comparison) and whether any of them
  is marked as trusted. A CIDR input only counts as contained when the whole range lies
  inside the location's range.
- Lists the Conditional Access policies that include or exclude those locations, plus
  policies that use 'AllTrusted' when the IP is in a trusted location.

With -Fix (supports -WhatIf):
  - With -LocationName: IPs that are not in a trusted location are added to that named
    location (as /32 or /128, or the network of the given CIDR), and the location is
    marked trusted. If the location does not exist it is created as trusted.
  - Without -LocationName: untrusted named locations that already contain the IP are
    marked trusted. Note: that trusts the complete range of that location.
  Afterwards the IPs are evaluated again.

Not covered: the legacy "trusted IPs" in the per-user MFA service settings.

Output: NamedLocations_<domain>_<date>.csv
```

## Parameters

- `-TenantId`: Tenant ID or initial domain.
- `-IpAddress`: One or more IP addresses or CIDR ranges, e.g. '203.0.113.10', '203.0.113.0/28', '2001:db8::/48'.
- `-ClientId`: App registration (client) ID for app-only authentication. Without it: interactive (delegated).
- `-CertificateThumbprint`: Certificate thumbprint for app-only authentication.
- `-Fix`: Repair (see description). Use with -WhatIf first.
- `-LocationName`: Named location to add the IPs to when using -Fix.
- `-SelfTest`: Runs the built-in CIDR tests locally, without connecting to a tenant.

## Examples

```powershell
.\Test-NamedLocationTrusted.ps1 -TenantId contoso.onmicrosoft.com -IpAddress 203.0.113.10
```

```powershell
.\Test-NamedLocationTrusted.ps1 -TenantId $tid -IpAddress 203.0.113.10,2001:db8:5::/48 -Fix -LocationName 'Kantoor' -WhatIf
```

```powershell
.\Test-NamedLocationTrusted.ps1 -SelfTest
```

## Notes

```text
Required Graph permissions (application or delegated):
  Policy.Read.All                       named locations and CA policies (check)
  Policy.ReadWrite.ConditionalAccess    only for -Fix
  Organization.Read.All                 tenant name (optional)
Delegated -Fix: the user also needs Conditional Access Administrator or Security Administrator.

Dependency: Microsoft.Graph.Authentication.
Exit codes: 0 = every IP is in a trusted named location, 1 = at least one is not,
2 = fatal error (or a failed self-test).
```
