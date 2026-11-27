# Get-ExpiringAppCredentials

Finds client secrets and certificates of app registrations (and optionally enterprise apps)
that expire within X days.

## Description

```text
Checks passwordCredentials (secrets) and keyCredentials (certificates) of:
  - app registrations (/applications)
  - with -IncludeServicePrincipals: enterprise apps (/servicePrincipals) of type
    'Application', which includes SAML signing certificates of gallery/SAML apps.

Reports credentials whose end date is within -Days (default 30). Already expired
credentials are included only with -IncludeExpired.

Column HasNewerCredential = True means the same object already has another credential
of the same type (and usage) that is valid beyond the window: the credential has probably
been rotated already and only the old one still needs to be removed.

Output: AppCredentials_<domain>_<date>.csv
```

## Parameters

- `-TenantId`: Tenant ID or initial domain.
- `-ClientId`: App registration (client) ID for app-only authentication. Without it: interactive (delegated).
- `-CertificateThumbprint`: Certificate thumbprint for app-only authentication.
- `-Days`: Window in days (default 30).
- `-IncludeExpired`: Also report credentials that have already expired.
- `-IncludeServicePrincipals`: Also check enterprise apps (service principals), e.g. SAML signing certificates.

## Examples

```powershell
.\Get-ExpiringAppCredentials.ps1 -TenantId contoso.onmicrosoft.com -Days 60 -IncludeServicePrincipals
```

## Notes

```text
Required Graph permissions (application or delegated):
  Application.Read.All    applications, servicePrincipals and their owners
  Organization.Read.All   tenant name (optional)
Owner UPNs of user owners may additionally need User.ReadBasic.All in some tenants.

Dependency: Microsoft.Graph.Authentication.
Exit codes: 0 = nothing expiring, 1 = expiring/expired credentials found, 2 = fatal error.
```
