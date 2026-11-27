# Users

Scripts for the user lifecycle in Microsoft 365: onboarding, changing roles, offboarding and cleaning up. They use Microsoft Graph and sign in the same way as the scripts in [Graph](../Graph): interactively, or app-only with `-ClientId` and `-CertificateThumbprint`.

The scripts that change something support `-WhatIf`. Run that first.

| Script | What it does |
|---|---|
| [New-M365User](New-M365User) | Creates a user with licence, manager and groups |
| [Copy-UserGroupMembership](Copy-UserGroupMembership) | Gives a user the same groups as a colleague |
| [Invoke-UserOffboarding](Invoke-UserOffboarding) | Blocks sign-in, ends sessions, removes groups and licences |
| [Get-InactiveUsers](Get-InactiveUsers) | Accounts without a sign-in for X days, with their licences |
| [Get-StaleGuestUsers](Get-StaleGuestUsers) | Guests that never accepted or no longer sign in |
