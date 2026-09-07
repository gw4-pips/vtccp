param([string]$Source = $PSScriptRoot)
$ErrorActionPreference = "Stop"
$target = Join-Path $env:LOCALAPPDATA "Programs\VCCS\FlexWedge"
New-Item -ItemType Directory -Force -Path $target | Out-Null
Copy-Item (Join-Path $Source "*") $target -Recurse -Force
$desktop = [Environment]::GetFolderPath("Desktop")
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut((Join-Path $desktop "VCCS FlexWedge.lnk"))
$shortcut.TargetPath = Join-Path $target "FlexWedge.exe"
$shortcut.WorkingDirectory = $target
$shortcut.Save()
Write-Host "Installed for current user: $target"