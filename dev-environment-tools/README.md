# Windows development environment audit

This folder is a safe, read-only inventory tool for the Windows laptop used
to build and test the VTCCP and GS1 tooling.

It does not install software, change PATH, modify registry settings, move
files, delete files, connect to devices, or change either product repository.

## Install the toolkit folder

Copy this entire folder to:

```text
C:\dev\dev-environment-tools
```

The scripts are written for both Windows PowerShell 5.1 and PowerShell 7.
They intentionally use ASCII-only `.ps1` files because Windows PowerShell 5.1
can misread non-ASCII script files.

## Run the complete audit

Open PowerShell and run:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
cd C:\dev\dev-environment-tools
.\Invoke-DevEnvironmentAudit.ps1 -DevRoot C:\dev
```

The default repository checks are:

```text
C:\dev\gs1-digital-link-resolver
C:\dev\vtccp
```

Command states are `FOUND` for a usable command, `BROKEN` when a command
resolves but its version check fails, and `MISSING` when no command is found.

The audit writes timestamped JSON and Markdown reports to:

```text
C:\dev\dev-environment-tools\reports
```

If you update the toolkit, replace the entire `dev-environment-tools` folder.
The entry script and the shared `DevEnvironment.Common.ps1` file must come from
the same bundle.

## Run focused checks

Toolchain only:

```powershell
.\Get-DevToolInventory.ps1
```

Repositories, vendor DLLs, WebView2 loader files, and wkhtmltopdf:

```powershell
.\Find-RepoAndSdkDependencies.ps1 -DevRoot C:\dev
```

Verify the VTCCP checkout and ASR SDK placement without changing anything:

```powershell
.\Confirm-VtccpSetup.ps1 -DevRoot C:\dev
```

After the repository and SDK DLL are in place, optionally run a build check:

```powershell
.\Confirm-VtccpSetup.ps1 -DevRoot C:\dev -RunBuild
```

This verifier never installs, moves, deletes, or overwrites files. The optional
build can create normal `bin` and `obj` build outputs.

## Run the home TC-829 / VeriWedge check

All Webscan TruCheck units are USB-connected devices. They do not receive a
network address and this check never opens a TCP connection to one. After the
VTCCP checkout is present, run the focused home-development check:

```powershell
.\Test-HomeVtccpVeriWedge.ps1 `
  -DevRoot C:\dev `
  -RunBuild `
  -RunValidationTests
```

The device portion inventories only Windows PnP USB entries named `Webscan` or
`TruCheck`; it sends no command and makes no setting change. A generic Windows
driver name produces a warning so the operator can confirm the device manually
in Device Manager. The validation tests exercise the bundled GS1 1.4.1 engine
against a known Digital Link, a GS1 Element String, and invalid input on Windows.
The script writes a timestamped JSON and Markdown result pair to `reports`.

### DataMan-only HTTP evidence capture

The following capture option belongs only to the DataMan verifier's existing
DMST/HTTP integration. It is not a Webscan result path and must not be used to
infer one. To preserve DataMan source evidence for a bench run, set this
process-only environment variable before launching the built app:

```powershell
$env:VTCCP_HTTP_CAPTURE_DIR = "$env:USERPROFILE\Documents\VTCCP-Diagnostic\TC-829"
& "C:\dev\vtccp\vtccp\VtccpApp\bin\Release\net8.0-windows10.0.18362.0\VtccpApp.exe"
```

For each DataMan HTTP verification result, VTCCP writes a paired
`pcm_report.html`, `codes.xml`, and decoded `push.xml` file to the chosen
folder. Treat these files as scan evidence: they can contain decoded barcode
data and should not be committed to Git. Leave the variable unset for normal
operation and do not set it for a Webscan run.

### Webscan result path

Webscan TruChecks use a separate result path from the DataMan DMST/HTTP
integration. This repository does not assume that path is HTTP, DMST, or the
DataMan `HttpEventSubscriber`. After the USB preflight passes, use the
Webscan-specific result/export workflow supplied for that installation; keep
the raw Webscan output and the VTCCP session output together for review. Do
not alter firmware, trigger type, reader settings, or the USB driver during
this test.

Use a different repository list:

```powershell
.\Invoke-DevEnvironmentAudit.ps1 `
  -DevRoot C:\dev `
  -RepoName @("gs1-digital-link-resolver", "another-repo")
```

## What the first audit is checking

- Git and optional GitHub CLI
- .NET SDKs and runtimes
- Visual Studio installations and workloads as exposed by `vswhere`
- MSBuild, PowerShell, Node.js, npm, pnpm, Python, 7-Zip, and Java
- WebView2 and the bundled `wkhtmltopdf.exe` fallback
- AsReader P35U SDK DLL candidates, including the documented `Q:` vendor path
- Repository existence, current branch, working-tree status, and remotes
- COM ports and filtered USB device entries for Cognex/DataMan/AsReader/RFID
- Installed program entries matching the tools relevant to this project

## Next step

Do not organize or install anything based on assumptions. Run the audit first
and bring back the Markdown report. We can then create a small explicit
machine layout and, only after review, a separate setup script for approved
folders, SDK placement, and build prerequisites.