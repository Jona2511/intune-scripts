# Compliance

Custom compliance checks for Windows. Each folder has a discovery script and a JSON file with the rules:

```
Compliance/
  <Name>/
    <Name>.ps1
    <Name>.json
```

The script returns its results as one line of compressed JSON and must output nothing else. The JSON file holds the rules Intune checks those results against, including the message users see when their device is not compliant. Each README explains the setup in Intune.

| Check | Compliant when |
|---|---|
| [Test-SophosAgent](Test-SophosAgent) | Sophos is installed and its services are running |
| [Test-HuntressAgent](Test-HuntressAgent) | Huntress is installed and its services are running |
| [Test-FirewallProfiles](Test-FirewallProfiles) | Windows Firewall is on for the domain, private and public profiles |
| [Test-ForbiddenApps](Test-ForbiddenApps) | No apps from the forbidden list are installed |
