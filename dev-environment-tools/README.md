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
integration. The confirmed TC-829 workflow is a local **Webscan HTML export**:

- The verifier is USB-connected. It has no TCP target, and VTCCP does not call
  `HttpEventSubscriber`, DMCC, or DataMan SDK code for a Webscan session.
- Webscan writes one `Webscan TruCheck™ USB Verification Report` HTML file and
  its sibling symbol image to:

  ```text
  C:\dev\vtccp\TC-829 VeriWedge Dev Reports
  ```

- In VTCCP, select a job template, choose **Webscan HTML (USB)** in Session
  Launcher, and start the session. A DataMan device profile is not required in
  this mode.
- Run the controlled verification from the Webscan software. When its new HTML
  report appears in the directory, VTCCP imports the literal report cells and
  appends one verification record to the active session workbook.
- VTCCP does not trigger the Webscan verifier, calculate grades, rewrite the
  report, delete the source HTML, or alter the sibling image. Keep the original
  HTML, image, and optional native PDF in that directory as the raw verifier
  evidence for the scan.

The controlled TC-829 report used to verify this contract has these literal
anchors: software `3.03.74`, serial `TC-829-0213-021`, DataMatrix,
ISO15415:2011, overall grade `A (4.0)`, aperture `08`, wavelength `660`, and
19 ISO15415 quality rows (including separate 3a Modulation and 3b Reflectance
Margin rows). It is covered by the VTCCP parser tests.

#### Limits and recovery

- This adapter accepts only new `.html` reports with the Webscan TruCheck title.
  PDF and `.xls` exports are evidence files, not live-import inputs.
- The report must finish writing before it can be imported. A failed parse is
  shown in the session status and leaves the source file untouched; correct the
  Webscan export and run another verification rather than editing the raw file.
- The watcher ignores duplicate file-system events for an unchanged export. If
  a report was created while VTCCP was not running, preserve it and create a
  fresh controlled verification after starting the Webscan session.
- Do not alter firmware, trigger type, reader settings, or the USB driver
  during this test.

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