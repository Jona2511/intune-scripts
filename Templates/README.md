# Templates

| File | Use it for |
|---|---|
| [Template-Script.ps1](Template-Script.ps1) | Platform scripts and anything else that runs once |
| [Template-Remediation.ps1](Template-Remediation.ps1) | Remediations: check with `-DetectOnly`, repair without it |

Both log to `%ProgramData%\<ScriptName>\<ScriptName>.log`, like the other scripts in this repo.

Copy the file, rename it, fill in the help block at the top and replace the placeholder code.
