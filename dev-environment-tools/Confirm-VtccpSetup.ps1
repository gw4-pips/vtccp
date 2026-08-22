[CmdletBinding()]
param(
    [string] $DevRoot = "C:\dev",
    [string] $RepoPath = "",
    [string] $AsReaderSource = "",
    [switch] $RunBuild,
    [switch] $AsObject
)

$ErrorActionPreference = "Continue"
. (Join-Path $PSScriptRoot "DevEnvironment.Common.ps1")

if ([string]::IsNullOrWhiteSpace($RepoPath)) {
    $RepoPath = Join-Path $DevRoot "vtccp"
}

$checks = @()

function New-SetupCheck {
    param(
        [string] $Name,
        [ValidateSet("PASS", "WARN", "FAIL")]
        [string] $Status,
        [string] $Detail
    )

    return [pscustomobject]@{
        Name   = $Name
        Status = $Status
        Detail = $Detail
    }
}

function Get-HashRecord {
    param([string] $Path)

    try {
        $hash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash
        return [pscustomobject]@{
            Path   = $Path
            Exists = $true
            Hash   = $hash
            Error  = $null
        }
    }
    catch {
        return [pscustomobject]@{
            Path   = $Path
            Exists = $true
            Hash   = $null
            Error  = $_.Exception.Message
        }
    }
}

$checks += New-SetupCheck `
    -Name "Development root" `
    -Status $(if (Test-Path -LiteralPath $DevRoot -PathType Container) { "PASS" } else { "FAIL" }) `
    -Detail $DevRoot

$solutionPath = Join-Path $RepoPath "vtccp\VTCCP.sln"
$expectedDllPath = Join-Path $RepoPath "lib\asreader-p3xu-sdk-1.3.0\AsReaderP3xU.dll"

if (-not (Test-Path -LiteralPath $RepoPath -PathType Container)) {
    $checks += New-SetupCheck -Name "VTCCP repository folder" -Status "FAIL" -Detail "$RepoPath is missing"
}
else {
    $checks += New-SetupCheck -Name "VTCCP repository folder" -Status "PASS" -Detail $RepoPath
}

$git = Get-Command -Name "git" -ErrorAction SilentlyContinue
$gitStatus = @()
if ($null -eq $git) {
    $checks += New-SetupCheck -Name "Git command" -Status "FAIL" -Detail "git was not found on PATH"
}
elseif (Test-Path -LiteralPath $RepoPath -PathType Container) {
    $gitRoot = @(Invoke-CapturedCommand -FilePath $git.Source -Arguments @(
        "-C", $RepoPath, "rev-parse", "--show-toplevel"
    )) | Where-Object { $_ -and ($_ -notmatch "^fatal:") } | Select-Object -First 1

    if ($gitRoot) {
        $checks += New-SetupCheck -Name "VTCCP Git checkout" -Status "PASS" -Detail $gitRoot
        $gitStatus = @(Invoke-CapturedCommand -FilePath $git.Source -Arguments @(
            "-C", $RepoPath, "status", "--short", "--branch"
        ))
        $dirtyLines = @($gitStatus | Where-Object {
            $_ -and ($_ -notmatch "^##")
        })
        if ($dirtyLines.Count -eq 0) {
            $checks += New-SetupCheck -Name "VTCCP working tree" -Status "PASS" -Detail "clean"
        }
        else {
            $checks += New-SetupCheck `
                -Name "VTCCP working tree" `
                -Status "WARN" `
                -Detail ("uncommitted entries: " + ($dirtyLines -join "; "))
        }
    }
    else {
        $checks += New-SetupCheck `
            -Name "VTCCP Git checkout" `
            -Status "FAIL" `
            -Detail "$RepoPath is not a Git checkout"
    }
}

if (Test-Path -LiteralPath $solutionPath -PathType Leaf) {
    $checks += New-SetupCheck -Name "VTCCP solution" -Status "PASS" -Detail $solutionPath
}
else {
    $checks += New-SetupCheck -Name "VTCCP solution" -Status "FAIL" -Detail "$solutionPath is missing"
}

$dotnet = Get-CommandInventory -Name "dotnet" -VersionArguments @("--version")
if (-not $dotnet.Available) {
    $checks += New-SetupCheck -Name ".NET SDK" -Status "FAIL" -Detail "dotnet was not found on PATH"
}
elseif (-not $dotnet.Healthy) {
    $checks += New-SetupCheck -Name ".NET SDK" -Status "FAIL" -Detail $dotnet.Version
}
else {
    $checks += New-SetupCheck -Name ".NET SDK" -Status "PASS" -Detail $dotnet.Version
}

$vsInventory = Get-VisualStudioInventory
$vsInstall = @($vsInventory.Installations |
    Where-Object { $_.InstallationPath } |
    Select-Object -First 1)
if ($vsInstall.Count -eq 0) {
    $checks += New-SetupCheck -Name "Visual Studio" -Status "FAIL" -Detail "no installation was found through vswhere"
}
else {
    $vsPath = $vsInstall[0].InstallationPath
    $vsName = if ($vsInstall[0].InstallationName) { $vsInstall[0].InstallationName } else { "Visual Studio" }
    $checks += New-SetupCheck -Name "Visual Studio" -Status "PASS" -Detail "$vsName at $vsPath"

    $msbuildPath = Join-Path $vsPath "MSBuild\Current\Bin\MSBuild.exe"
    if (Test-Path -LiteralPath $msbuildPath -PathType Leaf) {
        $checks += New-SetupCheck -Name "MSBuild" -Status "PASS" -Detail $msbuildPath
    }
    else {
        $checks += New-SetupCheck -Name "MSBuild" -Status "WARN" -Detail "$msbuildPath was not found"
    }
}

$sourceCandidates = @()
if (-not [string]::IsNullOrWhiteSpace($AsReaderSource)) {
    $sourceCandidates += $AsReaderSource
}
$sourceCandidates += @(
    (Join-Path $DevRoot "gs1-digital-link-resolver\tools\rfid-wedge\AsReaderP3xU.dll"),
    (Join-Path $env:USERPROFILE "Downloads\AsReader_P35U_SDK_cs_1_3_0\AsReader_P35U_SDK_c#_1_3_0\AsReaderP3xU.dll"),
    (Join-Path $env:USERPROFILE "Downloads\Sample_AsReader-P35U_cs_1_3_0(1)\Sample_AsReader-P35U_c#_1_3_0\AsReaderP3xU_Demo\bin\Release\AsReaderP3xU.dll")
)

$searchRoots = @(
    $DevRoot,
    (Join-Path $env:USERPROFILE "Downloads")
) | Where-Object {
    $_ -and (Test-Path -LiteralPath $_ -PathType Container)
}

foreach ($root in $searchRoots) {
    $sourceCandidates += @(Get-ChildItem -LiteralPath $root -Recurse -File -Filter "AsReaderP3xU.dll" -ErrorAction SilentlyContinue |
        ForEach-Object { $_.FullName })
}

$sourceCandidates = @($sourceCandidates |
    Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } |
    Sort-Object -Unique)

$hashRecords = @()
foreach ($candidate in $sourceCandidates) {
    $hashRecords += Get-HashRecord -Path $candidate
}

$expectedRecord = $null
if (Test-Path -LiteralPath $expectedDllPath -PathType Leaf) {
    $expectedRecord = Get-HashRecord -Path $expectedDllPath
    if ($expectedRecord.Hash) {
        $checks += New-SetupCheck -Name "VTCCP ASR DLL" -Status "PASS" -Detail "$expectedDllPath (SHA256 $($expectedRecord.Hash))"
    }
    else {
        $checks += New-SetupCheck -Name "VTCCP ASR DLL" -Status "FAIL" -Detail "$expectedDllPath could not be hashed: $($expectedRecord.Error)"
    }
}
else {
    $checks += New-SetupCheck -Name "VTCCP ASR DLL" -Status "FAIL" -Detail "$expectedDllPath is missing"
}

$sourceRecords = @($hashRecords | Where-Object {
    $_.Path -ne $expectedDllPath -and $_.Hash
})
if ($sourceRecords.Count -eq 0) {
    $checks += New-SetupCheck `
        -Name "ASR SDK source copy" `
        -Status "FAIL" `
        -Detail "No existing AsReaderP3xU.dll source was found"
}
else {
    $preferredSource = @($sourceRecords | Sort-Object `
        @{ Expression = {
            if ($_.Path -like "*AsReader_P35U_SDK_cs_1_3_0*") { 0 }
            elseif ($_.Path -like "*gs1-digital-link-resolver*") { 1 }
            else { 2 }
        } }, Path | Select-Object -First 1)
    $checks += New-SetupCheck `
        -Name "ASR SDK source copy" `
        -Status "PASS" `
        -Detail "$($preferredSource[0].Path) (SHA256 $($preferredSource[0].Hash))"

    if ($expectedRecord -and $expectedRecord.Hash -eq $preferredSource[0].Hash) {
        $checks += New-SetupCheck -Name "ASR DLL hash match" -Status "PASS" -Detail "project copy matches the selected source"
    }
    elseif ($expectedRecord) {
        $checks += New-SetupCheck -Name "ASR DLL hash match" -Status "FAIL" -Detail "project copy differs from the selected source"
    }
    else {
        $checks += New-SetupCheck -Name "ASR DLL hash match" -Status "WARN" -Detail "copy the selected source to the expected VTCCP path"
    }
}

if ($RunBuild) {
    if (-not (Test-Path -LiteralPath $solutionPath -PathType Leaf)) {
        $checks += New-SetupCheck -Name "VTCCP build" -Status "FAIL" -Detail "solution is missing"
    }
    elseif (-not $dotnet.Available) {
        $checks += New-SetupCheck -Name "VTCCP build" -Status "FAIL" -Detail "dotnet is unavailable"
    }
    else {
        $buildOutput = @(Invoke-CapturedCommand -FilePath $dotnet.Path -Arguments @(
            "build", $solutionPath, "--nologo"
        ))
        $buildSucceeded = @($buildOutput | Where-Object { $_ -match "Build succeeded" }).Count -gt 0
        if ($buildSucceeded) {
            $checks += New-SetupCheck -Name "VTCCP build" -Status "PASS" -Detail "dotnet build reported success"
        }
        else {
            $checks += New-SetupCheck -Name "VTCCP build" -Status "FAIL" -Detail (($buildOutput | Select-Object -Last 8) -join " | ")
        }
    }
}
else {
    $checks += New-SetupCheck -Name "VTCCP build" -Status "WARN" -Detail "not run; use -RunBuild after the DLL is in place"
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
    GeneratedAt       = (Get-Date).ToString("o")
    ComputerName      = $env:COMPUTERNAME
    DevRoot           = $DevRoot
    RepoPath          = $RepoPath
    SolutionPath      = $solutionPath
    ExpectedDllPath   = $expectedDllPath
    OverallStatus     = $overallStatus
    Checks            = $checks
    GitStatus         = $gitStatus
    AsReaderDllFiles  = $hashRecords
}

if ($AsObject) {
    return $result
}

Write-Host ""
Write-Host "VTCCP setup verification: $overallStatus"
foreach ($check in $checks) {
    Write-Host ("[{0}] {1}: {2}" -f $check.Status, $check.Name, $check.Detail)
}

$reportDir = Join-Path $PSScriptRoot "reports"
if (-not (Test-Path -LiteralPath $reportDir -PathType Container)) {
    New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
}
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$jsonPath = Join-Path $reportDir "vtccp-setup-$timestamp.json"
$mdPath = Join-Path $reportDir "vtccp-setup-$timestamp.md"

$result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonPath -Encoding UTF8

$markdown = @(
    "# VTCCP setup verification",
    "",
    "- Generated: $($result.GeneratedAt)",
    "- Computer: $($result.ComputerName)",
    "- Dev root: $($result.DevRoot)",
    "- Repository: $($result.RepoPath)",
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