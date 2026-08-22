[CmdletBinding()]
param(
    [string] $DevRoot = "C:\dev",
    [string[]] $RepoName = @("gs1-digital-link-resolver", "vtccp"),
    [string] $OutputDirectory = (Join-Path $PSScriptRoot "reports")
)

$ErrorActionPreference = "Continue"
. (Join-Path $PSScriptRoot "DevEnvironment.Common.ps1")

if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

$toolInventory = & (Join-Path $PSScriptRoot "Get-DevToolInventory.ps1") -AsObject
$repoInventory = & (Join-Path $PSScriptRoot "Find-RepoAndSdkDependencies.ps1") `
    -DevRoot $DevRoot -RepoName $RepoName -AsObject
$deviceInventory = Get-PnpInventory

$timestamp = Get-Date
$stamp = $timestamp.ToString("yyyyMMdd-HHmmss")
$jsonPath = Join-Path $OutputDirectory "dev-environment-audit-$stamp.json"
$markdownPath = Join-Path $OutputDirectory "dev-environment-audit-$stamp.md"

$audit = [pscustomobject]@{
    GeneratedAt = $timestamp.ToString("o")
    Computer    = $env:COMPUTERNAME
    User        = $env:USERNAME
    DevRoot     = $DevRoot
    Tools       = $toolInventory
    Repositories = $repoInventory.Repositories
    SdkCandidates = $repoInventory.SdkCandidates
    ExpectedPaths = $repoInventory.ExpectedPaths
    Devices     = $deviceInventory
}

$audit | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $jsonPath -Encoding UTF8

function Get-MarkdownToolLine {
    param([object] $Command)
    $state = if ($Command.Available) { "FOUND" } else { "MISSING" }
    $version = if ($Command.Version) { $Command.Version } else { "" }
    $path = if ($Command.Path) { $Command.Path } else { "" }
    return "| $($Command.Name) | $state | $version | $path |"
}

$md = New-Object System.Collections.Generic.List[string]
$md.Add("# Windows development environment audit")
$md.Add("")
$md.Add("- Generated: $($audit.GeneratedAt)")
$md.Add("- Computer: $($audit.Computer)")
$md.Add("- User: $($audit.User)")
$md.Add("- Dev root: $($audit.DevRoot)")
$md.Add("")
$md.Add("## Command inventory")
$md.Add("")
$md.Add("| Command | State | Version | Path |")
$md.Add("|---|---|---|---|")
foreach ($command in @($toolInventory.Commands)) {
    $md.Add((Get-MarkdownToolLine -Command $command))
}

$md.Add("")
$md.Add("## .NET SDKs and runtimes")
$md.Add("")
$md.Add("### SDKs")
foreach ($line in @($toolInventory.DotNet.Sdks)) {
    $md.Add("- $line")
}
$md.Add("")
$md.Add("### Runtimes")
foreach ($line in @($toolInventory.DotNet.Runtimes)) {
    $md.Add("- $line")
}

$md.Add("")
$md.Add("## Visual Studio")
$md.Add("")
foreach ($installation in @($toolInventory.VisualStudio.Installations)) {
    $md.Add("- $($installation.InstallationName) | $($installation.CatalogVersion) | $($installation.InstallationPath)")
}
if (@($toolInventory.VisualStudio.Installations).Count -eq 0) {
    $md.Add("- No Visual Studio installation was discovered through vswhere.")
}

$md.Add("")
$md.Add("## Repositories")
$md.Add("")
foreach ($repo in @($repoInventory.Repositories)) {
    $state = if ($repo.Exists) { "FOUND" } else { "MISSING" }
    $md.Add("### $($repo.Name): $state")
    $md.Add("- Path: $($repo.Path)")
    $md.Add("- Git root: $($repo.GitRoot)")
    $md.Add("- Branch: $($repo.Branch)")
    if (@($repo.Status).Count -gt 0) {
        $md.Add("- Git status:")
        foreach ($line in @($repo.Status)) {
            $md.Add("  - $line")
        }
    }
}

$md.Add("")
$md.Add("## Expected paths")
$md.Add("")
$md.Add("| Path | Exists |")
$md.Add("|---|---|")
foreach ($item in @($repoInventory.ExpectedPaths)) {
    $state = if ($item.Exists) { "YES" } else { "NO" }
    $md.Add("| $($item.Path) | $state |")
}

$md.Add("")
$md.Add("## SDK and runtime file candidates")
$md.Add("")
if (@($repoInventory.SdkCandidates).Count -eq 0) {
    $md.Add("- No AsReader, wkhtmltopdf, or WebView2Loader files were found in the scanned roots.")
}
else {
    foreach ($file in @($repoInventory.SdkCandidates)) {
        $md.Add("- $($file.FullName) | $($file.Length) bytes | $($file.LastWriteTime)")
    }
}

$md.Add("")
$md.Add("## Serial ports")
$md.Add("")
foreach ($port in @($deviceInventory.SerialPorts)) {
    $md.Add("- $($port.DeviceID) | $($port.Name) | $($port.Description) | $($port.Status)")
}
if (@($deviceInventory.SerialPorts).Count -eq 0) {
    $md.Add("- No serial ports were returned.")
}

$md.Add("")
$md.Add("## Filtered USB devices")
$md.Add("")
foreach ($device in @($deviceInventory.PnpUsb)) {
    $md.Add("- $($device.Status) | $($device.FriendlyName) | $($device.InstanceId)")
}
if (@($deviceInventory.PnpUsb).Count -eq 0) {
    $md.Add("- No matching Cognex, AsReader, RFID, serial, or USB device names were returned.")
}

$md | Set-Content -LiteralPath $markdownPath -Encoding UTF8

Write-Output "Audit complete."
Write-Output "JSON: $jsonPath"
Write-Output "Markdown: $markdownPath"
