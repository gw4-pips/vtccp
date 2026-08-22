[CmdletBinding()]
param(
    [string] $DevRoot = "C:\dev",
    [string[]] $RepoName = @("gs1-digital-link-resolver", "vtccp"),
    [switch] $AsObject
)

$ErrorActionPreference = "Continue"
. (Join-Path $PSScriptRoot "DevEnvironment.Common.ps1")

function Get-GitRepositoryInventory {
    param([string] $Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        return [pscustomobject]@{
            Name       = Split-Path -Leaf $Path
            Path       = $Path
            Exists     = $false
            GitRoot    = $null
            Branch     = $null
            Status     = @()
            Remotes    = @()
        }
    }

    $git = Get-Command -Name "git" -ErrorAction SilentlyContinue
    $gitRoot = $null
    $branch = $null
    $status = @()
    $remotes = @()

    if ($null -ne $git) {
        $gitRootLines = @(Invoke-CapturedCommand -FilePath $git.Source -Arguments @(
            "-C", $Path, "rev-parse", "--show-toplevel"
        ))
        $gitRoot = $gitRootLines |
            Where-Object { $_ -and ($_ -notmatch "^fatal:") } |
            Select-Object -First 1

        $branchLines = @(Invoke-CapturedCommand -FilePath $git.Source -Arguments @(
            "-C", $Path, "branch", "--show-current"
        ))
        $branch = $branchLines | Select-Object -First 1

        $status = @(Invoke-CapturedCommand -FilePath $git.Source -Arguments @(
            "-C", $Path, "status", "--short", "--branch"
        ))
        $remotes = @(Invoke-CapturedCommand -FilePath $git.Source -Arguments @(
            "-C", $Path, "remote", "-v"
        ))
    }

    return [pscustomobject]@{
        Name       = Split-Path -Leaf $Path
        Path       = $Path
        Exists     = $true
        GitRoot    = $gitRoot
        Branch     = $branch
        Status     = $status
        Remotes    = $remotes
    }
}

$repositories = @()
foreach ($name in $RepoName) {
    $repositories += Get-GitRepositoryInventory -Path (Join-Path $DevRoot $name)
}

$candidateRoots = @(
    $DevRoot,
    (Join-Path $DevRoot "vtccp"),
    "Q:\VendorDOC",
    (Join-Path $env:USERPROFILE "Downloads"),
    (Join-Path $env:USERPROFILE "Desktop")
) | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Container) } |
    Select-Object -Unique

$sdkFiles = @(Get-RelativeFileInventory -Root $DevRoot -Names @(
    "AsReaderP3xU.dll",
    "AsReaderP3xU*.dll",
    "wkhtmltopdf.exe",
    "WebView2Loader.dll"
))

foreach ($root in $candidateRoots) {
    if ($root -eq $DevRoot) {
        continue
    }
    $sdkFiles += @(Get-RelativeFileInventory -Root $root -Names @(
        "AsReaderP3xU.dll",
        "AsReaderP3xU*.dll",
        "wkhtmltopdf.exe",
        "WebView2Loader.dll"
    ))
}
$sdkFiles = @($sdkFiles | Sort-Object FullName -Unique)

$expectedPaths = @(
    (Join-Path $DevRoot "vtccp"),
    (Join-Path $DevRoot "vtccp\vtccp"),
    (Join-Path $DevRoot "vtccp\vtccp\VTCCP.sln"),
    (Join-Path $DevRoot "vtccp\lib\asreader-p3xu-sdk-1.3.0\AsReaderP3xU.dll"),
    "Q:\VendorDOC\AsReader (Asterisk)\AsReader_P35U_SDK_cs_1_3_0\AsReader_P35U_SDK_c#_1_3_0\AsReaderP3xU.dll"
) | ForEach-Object {
    [pscustomobject]@{
        Path   = $_
        Exists = Test-Path -LiteralPath $_
    }
}

$result = [pscustomobject]@{
    DevRoot       = $DevRoot
    DevRootExists = Test-Path -LiteralPath $DevRoot -PathType Container
    Repositories  = $repositories
    SdkCandidates = $sdkFiles
    ExpectedPaths = $expectedPaths
}

if ($AsObject) {
    return $result
}

$result | ConvertTo-Json -Depth 8
