[CmdletBinding()]
param(
    [switch] $AsObject
)

$ErrorActionPreference = "Continue"
. (Join-Path $PSScriptRoot "DevEnvironment.Common.ps1")

$commandSpecs = @(
    @{ Name = "git";       Args = @("--version") },
    @{ Name = "gh";        Args = @("--version") },
    @{ Name = "dotnet";    Args = @("--version") },
    @{ Name = "msbuild";   Args = @("-version") },
    @{ Name = "devenv";    Args = @("/version") },
    @{ Name = "vswhere";   Args = @("-version") },
    @{ Name = "powershell";Args = @("-NoProfile", "-Command", '$PSVersionTable.PSVersion.ToString()') },
    @{ Name = "pwsh";      Args = @("--version") },
    @{ Name = "node";      Args = @("--version") },
    @{ Name = "npm";       Args = @("--version") },
    @{ Name = "pnpm";      Args = @("--version") },
    @{ Name = "python";    Args = @("--version") },
    @{ Name = "py";        Args = @("--version") },
    @{ Name = "7z";        Args = @() },
    @{ Name = "java";      Args = @("-version") }
)

$commands = @()
foreach ($spec in $commandSpecs) {
    $commands += Get-CommandInventory -Name $spec.Name -VersionArguments $spec.Args
}

$result = [pscustomobject]@{
    ComputerName       = $env:COMPUTERNAME
    UserName           = $env:USERNAME
    PowerShellVersion  = $PSVersionTable.PSVersion.ToString()
    Is64BitOperatingSystem = [Environment]::Is64BitOperatingSystem
    Is64BitProcess     = [Environment]::Is64BitProcess
    Commands           = $commands
    DotNet             = Get-DotNetInventory
    VisualStudio       = Get-VisualStudioInventory
    InstalledPrograms  = @(Get-InstalledProgramInventory)
    PathEntries        = @(Get-PathInventory)
}

if ($AsObject) {
    return $result
}

$result | ConvertTo-Json -Depth 8
