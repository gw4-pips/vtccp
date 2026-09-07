param([switch]$PurgeUserData)
$ErrorActionPreference = "Stop"
$target = Join-Path $env:LOCALAPPDATA "Programs\VCCS\FlexWedge"
$userData = Join-Path $env:LOCALAPPDATA "VCCS\FlexWedge"
$shortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "VCCS FlexWedge.lnk"
Remove-Item $shortcut -Force -ErrorAction SilentlyContinue
Remove-Item $target -Recurse -Force -ErrorAction SilentlyContinue
if ($PurgeUserData) {
    Remove-Item $userData -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "FlexWedge program and user data removed."
    exit
}
Write-Host "FlexWedge removed for current user."
Write-Host "Configuration and session data preserved at: $userData"