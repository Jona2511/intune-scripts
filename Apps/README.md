# Apps

One folder per app, packaged as a Win32 app with the [Microsoft Win32 Content Prep Tool](https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool).

```
Apps/
  <AppName>/
    Install.ps1
    Uninstall.ps1
    Detect.ps1
    README.md
```

The README in each app folder lists the install and uninstall commands to enter in Intune, the detection rule and anything the app needs first. Installers and `.intunewin` files are not in the repo.
