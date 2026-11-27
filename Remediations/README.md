# Remediations

Scripts for Intune Remediations (Devices > Scripts and remediations). Most of them do detection and repair in one file: with `-DetectOnly` they only check, exiting with 1 when the device needs fixing and 0 when it doesn't, and without it they repair. Intune runs script files without arguments, so for the detection script upload a copy with `-DetectOnly` set as the default.

Two work differently. `Invoke-LocalAdminCheck` only reports by default and removes members when you add `-Remove`. `Invoke-LowDiskCleanup` has no detection mode: it cleans up when free space is below the threshold and exits with 1 if it is still too low afterwards.

| Script | What it does |
|---|---|
| [Invoke-BitLockerEscrowCheck](Invoke-BitLockerEscrowCheck) | Checks BitLocker status and makes sure recovery keys are escrowed to Entra ID |
| [Invoke-LocalAdminCheck](Invoke-LocalAdminCheck) | Finds unexpected members of the local Administrators group and reports or removes them |
| [Invoke-LowDiskCleanup](Invoke-LowDiskCleanup) | Cleans up C: when free space drops below a threshold |
| [Invoke-OneDrivePerMachineCheck](Invoke-OneDrivePerMachineCheck) | Moves per-user OneDrive installs to the per-machine client |
| [Invoke-TimeSyncCheck](Invoke-TimeSyncCheck) | Checks the Windows Time service and clock drift, and repairs them |
| [Invoke-UptimeWarning](Invoke-UptimeWarning) | Shows the user a reboot reminder when the device has been up too long |
| [Invoke-WindowsUpdateRepair](Invoke-WindowsUpdateRepair) | Detects a stuck Windows Update and resets the update components |
