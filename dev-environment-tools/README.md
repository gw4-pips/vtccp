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