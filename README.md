# Remove-MassLynx

I made this tool because we hit problems uninstalling MassLynx through the official Windows menu /
the official uninstaller: it ran almost to the end, rolled itself back, and left MassLynx installed.

## Why that happens

`Waters MassLynx V4.2 SCN1051.msi` is an InstallShield chainer: it installs the data packages and the
instrument driver inside one transaction, and at the very end of an uninstall it still has to run a
few InstallScript actions, unregister the `MassLynxSecurity` .NET assembly and clean up files. If any
of that fails, the whole transaction is rolled back. A running `mlynx4.exe` or `LancasterUI.exe`, or a
configured spectrometer that is not attached, has the same effect.

This tool uninstalls each product separately with its own `msiexec /x` and its own log, removes what
is left behind, and then checks whether it is really gone.

## Run it

In cmd or PowerShell — no git, no GitHub CLI, no token:

```bat
powershell -NoProfile -ExecutionPolicy Bypass -Command "$d=Join-Path $env:USERPROFILE 'masslynx-removal'; New-Item -ItemType Directory -Force $d | Out-Null; foreach($f in @('Start-MassLynxRemoval.ps1','Remove-MassLynx.ps1','Remove-MassLynx.cmd')){ Invoke-WebRequest -Uri ('https://raw.githubusercontent.com/MRM15R/masslynx-removal-tool/main/'+$f) -OutFile (Join-Path $d $f) -UseBasicParsing }; & (Join-Path $d 'Start-MassLynxRemoval.ps1')"
```

It downloads the three files into `%USERPROFILE%\masslynx-removal`, stages them in
`C:\ProgramData\MassLynxOnderhoud\app` and starts the tool; Windows asks for administrator rights.
Or download the repository as a ZIP and double-click `Remove-MassLynx.cmd`.

## What it does

1. Closes MassLynx and stops the services that live under `C:\MassLynx`.
2. Optionally copies `C:\MassLynx` to `C:\MassLynx-backup-<date>-<time>` first.
3. Uninstalls each MassLynx product on its own, driver and data packages first, the main product last.
4. Removes the leftovers: `C:\MassLynx`, `C:\TargetLynx`, COM registrations, shortcuts, the security
   assembly and the MassLynx registry values. `Waters Instruments` and the `Waters` key only when
   they end up empty.
5. Verifies the result with ten checks, including the Windows Installer database.

It does not install anything.

## What it leaves alone

Everything that is not MassLynx. Removal is matched on the MassLynx product codes, the MassLynx names
and the executables under `C:\MassLynx` — not on the word "Waters". Other Waters software is recorded
before the removal and checked afterwards; if one of them disappeared, the tool reports it.

## Logs

`C:\ProgramData\MassLynxOnderhoud\Logs\MassLynx-verwijderen-<date>-<time>.log`, plus the raw
Windows Installer log per product. On an error the log also holds the message, the script line and the
stack trace.

## Requirements

Windows 10 / 11 / Server 2016 or newer, PowerShell 5.1, local administrator rights.

## Licence

All rights reserved; see [LICENSE](LICENSE).
