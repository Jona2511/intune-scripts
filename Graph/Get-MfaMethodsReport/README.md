# Get-MfaMethodsReport

Reports users without MFA and users whose only MFA methods are phone-based (SMS/voice).

## Description

```text
Source: /reports/authenticationMethods/userRegistrationDetails (v1.0). Microsoft refreshes
this data periodically, so it is not real-time.

Categories:
  NoMfa      isMfaRegistered = false
  PhoneOnly  MFA registered, but every MFA method is a phone method
             (mobilePhone, alternateMobilePhone, officePhone, sms, voice).
             email, securityQuestion and temporaryAccessPass do not count as MFA methods.

Account status (enabled/disabled) and last sign-in come from /users. Last sign-in
(signInActivity) needs Entra ID P1/P2 and AuditLog.Read.All; if it is not available the
script continues without it.

Output: Mfa_Findings_<domain>_<date>.csv (and with -ExportAll also Mfa_AllUsers_...).
```

## Parameters

- `-TenantId`: Tenant ID or initial domain.
- `-ClientId`: App registration (client) ID for app-only authentication. Without it: interactive (delegated).
- `-CertificateThumbprint`: Certificate thumbprint for app-only authentication.
- `-ExcludeGuests`: Leave guest accounts out.
- `-ExcludeDisabled`: Leave disabled accounts out.
- `-SkipSignInActivity`: Do not request signInActivity (e.g. for tenants without P1).
- `-ExportAll`: Also export all users with their registration details.

## Examples

```powershell
.\Get-MfaMethodsReport.ps1 -TenantId contoso.onmicrosoft.com -ExcludeGuests -ExcludeDisabled
```

## Notes

```text
Required Graph permissions (application or delegated):
  AuditLog.Read.All                  userRegistrationDetails and signInActivity
  UserAuthenticationMethod.Read.All  userRegistrationDetails
  User.Read.All                      accountEnabled, userType
  Organization.Read.All              tenant name (optional)
Delegated: the user also needs e.g. Reports Reader, Security Reader or Global Reader.
Licensing: the authentication methods registration report requires Entra ID P1/P2.

Dependency: Microsoft.Graph.Authentication.
Exit codes: 0 = no findings, 1 = findings, 2 = fatal error.
```
