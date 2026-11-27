# Apps

Win32 apps, packaged with the [Microsoft Win32 Content Prep Tool](https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool). One script handles install, uninstall and detection:

- Install command: the script with its parameters
- Uninstall command: the same script with `-Uninstall`
- Detection: a custom detection script, which is a copy of the script with `-DetectOnly` set as the default

Each app's README has the exact commands to enter in Intune. Installers and `.intunewin` files are not in the repo.

| Script | What it does |
|---|---|
| [Invoke-WingetApp](Invoke-WingetApp) | Installs, upgrades, uninstalls or detects any winget app as SYSTEM |
| [Invoke-XelionInstall](Invoke-XelionInstall) | Silent install, uninstall and detection for the Xelion softphone |
