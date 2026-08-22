[CmdletBinding()]
param(
    [string] $DevRoot = "C:\dev",
    [string] $RepoPath = "",
    [switch] $RunBuild,
    [switch] $RunValidationTests,
    [switch] $AsObject
)

$ErrorActionPreference = "Continue"
. (Join-Path $PSScriptRoot "DevEnvironment.Common.ps1")

if ([string]::IsNullOrWhiteSpace($RepoPath)) {
    $RepoPath = Join-Path $DevRoot "vtccp"
}

function New-HomeCheck {
    param(
        [string] $Name,
        [ValidateSet("PASS", "WARN", "FAIL")]
        [string] $Status,
        [string] $Detail
    )

    return [pscustomobject]@{
        Name = $Name
        Status = $Status
        Detail = $Detail
    }
}

$checks = @()
$solutionPath = Join-Path $RepoPath "vtccp\VTCCP.sln"
$testProjectPath = Join-Path $RepoPath "vtccp\DeviceInterface.Tests\DeviceInterface.Tests.csproj"
$versionPinPath = Join-Path $RepoPath "vtccp\lib\gs1-syntax-engine\VERSION-PIN.md"
$validationPath = Join-Path $RepoPath "vtccp\DeviceInterface\Validation\VccsDigitalLinkValidationService.cs"
$nativeEnginePath = Join-Path $RepoPath "vtccp\lib\gs1-syntax-engine\dotnet-lib-release\runtimes\win-x64\native\gs1encoders.dll"
$dictionaryPath = Join-Path $RepoPath "vtccp\lib\gs1-syntax-engine\dotnet-lib-release\gs1-syntax-dictionary.txt"

if (Test-Path -LiteralPath $solutionPath -PathType Leaf) {
    $checks += New-HomeCheck -Name "VTCCP solution" -Status "PASS" -Detail $solutionPath
}
else {
    $checks += New-HomeCheck -Name "VTCCP solution" -Status "FAIL" -Detail "$solutionPath is missing"
}

if (Test-Path -LiteralPath $versionPinPath -PathType Leaf) {
    $pinMatches = @(Select-String -LiteralPath $versionPinPath -Pattern "Pinned version.*1\.4\.1")
    if ($pinMatches.Count -gt 0) {
        $checks += New-HomeCheck -Name "GS1 engine pin" -Status "PASS" -Detail "Bundled pin is 1.4.1"
    }
    else {
        $checks += New-HomeCheck -Name "GS1 engine pin" -Status "FAIL" -Detail "VERSION-PIN.md does not declare 1.4.1"
    }
}
else {
    $checks += New-HomeCheck -Name "GS1 engine pin" -Status "FAIL" -Detail "$versionPinPath is missing"
}

if (Test-Path -LiteralPath $validationPath -PathType Leaf) {
    $labelMatches = @(Select-String -LiteralPath $validationPath -Pattern 'EngineVersion = "GS1 Barcode Syntax Engine 1\.4\.1"')
    if ($labelMatches.Count -gt 0) {
        $checks += New-HomeCheck -Name "GS1 engine provenance label" -Status "PASS" -Detail "Validation source reports 1.4.1"
    }
    else {
        $checks += New-HomeCheck -Name "GS1 engine provenance label" -Status "FAIL" -Detail "Validation source does not report 1.4.1"
    }
}
else {
    $checks += New-HomeCheck -Name "GS1 engine provenance label" -Status "FAIL" -Detail "$validationPath is missing"
}

foreach ($asset in @(
    [pscustomobject]@{ Name = "GS1 native engine"; Path = $nativeEnginePath },
    [pscustomobject]@{ Name = "GS1 syntax dictionary"; Path = $dictionaryPath }
)) {
    if (Test-Path -LiteralPath $asset.Path -PathType Leaf) {
        $checks += New-HomeCheck -Name $asset.Name -Status "PASS" -Detail $asset.Path
    }
    else {
        $checks += New-HomeCheck -Name $asset.Name -Status "FAIL" -Detail "$($asset.Path) is missing"
    }
}

$pnpInventory = Get-PnpInventory
$webscanUsb = @($pnpInventory.PnpUsb | Where-Object {
    $_.FriendlyName -match "Webscan|TruCheck"
})
$readyWebscanUsb = @($webscanUsb | Where-Object { $_.Status -eq "OK" })

if ($readyWebscanUsb.Count -gt 0) {
    $deviceDetail = @($readyWebscanUsb | ForEach-Object {
        "$($_.FriendlyName) [$($_.Status)]"
    }) -join "; "
    $checks += New-HomeCheck `
        -Name "Webscan TruCheck USB presence" `
        -Status "PASS" `
        -Detail $deviceDetail
}
else {
    $deviceDetail = if ($webscanUsb.Count -gt 0) {
        @($webscanUsb | ForEach-Object {
            "$($_.FriendlyName) [$($_.Status)]"
        }) -join "; "
    }
    else {
        "No PnP USB entry named Webscan or TruCheck was found. Confirm the USB cable, powered device, and Windows Device Manager entry."
    }
    $checks += New-HomeCheck `
        -Name "Webscan TruCheck USB presence" `
        -Status "WARN" `
        -Detail $deviceDetail
}

$dotnet = Get-CommandInventory -Name "dotnet" -VersionArguments @("--version")
if (-not $dotnet.Available -or -not $dotnet.Healthy) {
    $checks += New-HomeCheck -Name ".NET SDK" -Status "FAIL" -Detail "dotnet is not ready: $($dotnet.Version)"
}
else {
    $checks += New-HomeCheck -Name ".NET SDK" -Status "PASS" -Detail $dotnet.Version
}

if ($RunBuild) {
    if (-not (Test-Path -LiteralPath $solutionPath -PathType Leaf) -or -not $dotnet.Healthy) {
        $checks += New-HomeCheck -Name "VTCCP build" -Status "FAIL" -Detail "Solution or dotnet SDK is unavailable."
    }
    else {
        $buildOutput = @(Invoke-CapturedCommand -FilePath $dotnet.Path -Arguments @(
            "build", $solutionPath, "-c", "Release", "--nologo"
        ))
        $buildPassed = @($buildOutput | Where-Object { $_ -match "Build succeeded" }).Count -gt 0
        $checks += New-HomeCheck `
            -Name "VTCCP build" `
            -Status $(if ($buildPassed) { "PASS" } else { "FAIL" }) `
            -Detail $(if ($buildPassed) { "Release build succeeded." } else { ($buildOutput | Select-Object -Last 10) -join " | " })
    }
}
else {
    $checks += New-HomeCheck -Name "VTCCP build" -Status "WARN" -Detail "Not run. Add -RunBuild to create normal bin/obj outputs."
}

if ($RunValidationTests) {
    if (-not (Test-Path -LiteralPath $testProjectPath -PathType Leaf) -or -not $dotnet.Healthy) {
        $checks += New-HomeCheck -Name "VeriWedge validation tests" -Status "FAIL" -Detail "Test project or dotnet SDK is unavailable."
    }
    else {
        $testOutput = @(Invoke-CapturedCommand -FilePath $dotnet.Path -Arguments @(
            "test", $testProjectPath, "-c", "Release", "--no-restore",
            "--filter", "FullyQualifiedName~VccsDigitalLinkValidationServiceTests",
            "--nologo"
        ))
        $testsPassed = @($testOutput | Where-Object { $_ -match "Passed!" }).Count -gt 0
        $checks += New-HomeCheck `
            -Name "VeriWedge validation tests" `
            -Status $(if ($testsPassed) { "PASS" } else { "FAIL" }) `
            -Detail $(if ($testsPassed) { "Digital Link, Element String, and invalid-input checks passed." } else { ($testOutput | Select-Object -Last 12) -join " | " })
    }
}
else {
    $checks += New-HomeCheck `
        -Name "VeriWedge validation tests" `
        -Status "WARN" `
        -Detail "Not run. Add -RunValidationTests after the checkout has restored packages."
}

$overallStatus = if (@($checks | Where-Object { $_.Status -eq "FAIL" }).Count -gt 0) {
    "FAIL"
}
elseif (@($checks | Where-Object { $_.Status -eq "WARN" }).Count -gt 0) {
    "WARN"
}
else {
    "PASS"
}

$result = [pscustomobject]@{
    GeneratedAt = (Get-Date).ToString("o")
    RepoPath = $RepoPath
    DeviceTransport = "USB"
    WebscanUsbCandidates = $webscanUsb
    OverallStatus = $overallStatus
    Checks = $checks
}

if ($AsObject) {
    return $result
}

Write-Host ""
Write-Host "Home VTCCP VeriWedge check: $overallStatus"
foreach ($check in $checks) {
    Write-Host ("[{0}] {1}: {2}" -f $check.Status, $check.Name, $check.Detail)
}

$reportDir = Join-Path $PSScriptRoot "reports"
if (-not (Test-Path -LiteralPath $reportDir -PathType Container)) {
    New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$jsonPath = Join-Path $reportDir "home-vtccp-veriwedge-$timestamp.json"
$mdPath = Join-Path $reportDir "home-vtccp-veriwedge-$timestamp.md"
$result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonPath -Encoding UTF8

$markdown = @(
    "# Home VTCCP VeriWedge check",
    "",
    "- Generated: $($result.GeneratedAt)",
    "- Repository: $($result.RepoPath)",
    "- Device transport: USB (Webscan TruCheck; no network endpoint is used)",
    "- Overall status: **$($result.OverallStatus)**",
    "",
    "| Status | Check | Detail |",
    "|---|---|---|"
)
foreach ($check in $checks) {
    $safeDetail = ($check.Detail -replace "\|", "\|")
    $markdown += "| $($check.Status) | $($check.Name) | $safeDetail |"
}
Set-Content -LiteralPath $mdPath -Value $markdown -Encoding UTF8

Write-Host ""
Write-Host "JSON: $jsonPath"
Write-Host "Markdown: $mdPath"