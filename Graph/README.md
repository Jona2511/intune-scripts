# Graph

Scripts that use Microsoft Graph for reporting and bulk changes.

Each script lists the permissions it needs in its README. Without `-ClientId` a script signs in interactively. For unattended runs, pass `-ClientId` and `-CertificateThumbprint` for an app registration with a certificate. None of the scripts use or store client secrets.

| Script | What it does |
|---|---|
| [Get-StaleDevices](Get-StaleDevices) | Intune and Entra ID devices that have not checked in for more than X days |
| [Get-LicenseOverview](Get-LicenseOverview) | Subscriptions, assignments per user and users without a licence |
| [Get-MfaMethodsReport](Get-MfaMethodsReport) | Users without MFA, or with only SMS or voice |
| [Get-ExpiringAppCredentials](Get-ExpiringAppCredentials) | Client secrets and certificates of app registrations that are about to expire |
| [Set-IntuneAssignments](Set-IntuneAssignments) | Assigns Intune apps and policies to groups in bulk |
| [Test-NamedLocationTrusted](Test-NamedLocationTrusted) | Checks whether IP addresses fall inside a trusted named location for Conditional Access |
