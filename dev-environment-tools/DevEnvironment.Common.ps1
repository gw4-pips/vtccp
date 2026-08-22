# Common helpers for the Windows development-environment audit.
# This file is intentionally compatible with Windows PowerShell 5.1.

Set-StrictMode -Version 2.0

function Invoke-CapturedCommand {
    param(
        [Parameter(Mandatory = $true)]
        [string] $FilePath,

        [string[]] $Arguments = @()
    )

    $lines = @()
    try {
        $lines = @(& $FilePath @Arguments 2>&1 | ForEach-Object { "$_" })
    }
    catch {
        $lines = @("ERROR: $($_.Exception.Message)")
    }
    return $lines
}

function Get-OptionalPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Object,

        [Parameter(Mandatory = $true)]
        [string] $Name
    )

    if ($null -eq $Object) {
        return $null
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Get-CommandInventory {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Name,

        [string[]] $VersionArguments = @("--version")
    )

    $command = Get-Command -Name $Name -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandType -in @("Application", "Cmdlet", "Function") } |
        Select-Object -First 1

    if ($null -eq $command) {
        return [pscustomobject]@{
            Name      = $Name
            Available = $false
            Path      = $null
            Version   = $null
            Details   = @()
        }
    }

    $path = $command.Source
    if ([string]::IsNullOrWhiteSpace($path)) {
        $path = $command.Definition
    }

    $details = @(Invoke-CapturedCommand -FilePath $path -Arguments $VersionArguments)
    $version = $details |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -First 1

    return [pscustomobject]@{
        Name      = $Name
        Available = $true
        Path      = $path
        Version   = $version
        Details   = $details
    }
}

function Get-DotNetInventory {
    $dotnet = Get-CommandInventory -Name "dotnet" -VersionArguments @("--version")
    if (-not $dotnet.Available) {
        return [pscustomobject]@{
            Available = $false
            Command   = $dotnet
            Sdks      = @()
            Runtimes  = @()
        }
    }

    $sdks = @(Invoke-CapturedCommand -FilePath $dotnet.Path -Arguments @("--list-sdks"))
    $runtimes = @(Invoke-CapturedCommand -FilePath $dotnet.Path -Arguments @("--list-runtimes"))

    return [pscustomobject]@{
        Available = $true
        Command   = $dotnet
        Sdks      = $sdks
        Runtimes  = $runtimes
    }
}

function Get-VisualStudioInventory {
    $candidates = @()
    if ($env:ProgramFiles) {
        $candidates += (Join-Path $env:ProgramFiles "Microsoft Visual Studio\Installer\vswhere.exe")
    }
    if (${env:ProgramFiles(x86)}) {
        $candidates += (Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe")
    }

    $vswhere = $candidates |
        Where-Object { Test-Path -LiteralPath $_ } |
        Select-Object -First 1

    if ($null -eq $vswhere) {
        return [pscustomobject]@{
            Available       = $false
            VsWherePath     = $null
            Installations   = @()
            RawDetails      = @()
        }
    }

    $raw = @(Invoke-CapturedCommand -FilePath $vswhere -Arguments @(
        "-all", "-products", "*", "-format", "json"
    ))
    $installations = @()

    try {
        $parsed = ($raw -join "`n") | ConvertFrom-Json
        foreach ($item in @($parsed)) {
            $catalogInfo = Get-OptionalPropertyValue -Object $item -Name "catalogInfo"
            $installations += [pscustomobject]@{
                InstallationPath = Get-OptionalPropertyValue -Object $item -Name "installationPath"
                InstallationName = Get-OptionalPropertyValue -Object $item -Name "displayName"
                ProductId        = Get-OptionalPropertyValue -Object $item -Name "productId"
                CatalogVersion   = Get-OptionalPropertyValue -Object $catalogInfo -Name "productDisplayVersion"
                IsComplete       = Get-OptionalPropertyValue -Object $item -Name "isComplete"
                IsLaunchable     = Get-OptionalPropertyValue -Object $item -Name "isLaunchable"
                ParseError       = $null
                RawOutput        = @()
            }
        }
    }
    catch {
        $installations = @([pscustomobject]@{
            InstallationPath = $null
            InstallationName = $null
            ProductId        = $null
            CatalogVersion   = $null
            IsComplete       = $null
            IsLaunchable     = $null
            ParseError       = $_.Exception.Message
            RawOutput        = $raw
        })
    }

    return [pscustomobject]@{
        Available     = $true
        VsWherePath   = $vswhere
        Installations = $installations
        RawDetails    = $raw
    }
}

function Get-InstalledProgramInventory {
    $roots = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $terms = @(
        "Visual Studio", ".NET", "WebView2", "Cognex", "DataMan",
        "TruCheck", "Webscan", "AsReader", "Node.js", "Git"
    )
    $matches = @()

    foreach ($root in $roots) {
        $items = @(Get-ItemProperty -Path $root -ErrorAction SilentlyContinue)
        foreach ($item in $items) {
            $displayName = Get-OptionalPropertyValue -Object $item -Name "DisplayName"
            if ([string]::IsNullOrWhiteSpace($displayName)) {
                continue
            }
            $isMatch = $false
            foreach ($term in $terms) {
                if ($displayName -like "*$term*") {
                    $isMatch = $true
                    break
                }
            }
            if ($isMatch) {
                $matches += [pscustomobject]@{
                    DisplayName     = $displayName
                    DisplayVersion  = Get-OptionalPropertyValue -Object $item -Name "DisplayVersion"
                    Publisher       = Get-OptionalPropertyValue -Object $item -Name "Publisher"
                    InstallLocation = Get-OptionalPropertyValue -Object $item -Name "InstallLocation"
                    UninstallKey    = Get-OptionalPropertyValue -Object $item -Name "PSPath"
                }
            }
        }
    }

    return @($matches | Sort-Object DisplayName, DisplayVersion -Unique)
}

function Get-PnpInventory {
    $ports = @()
    $usb = @()

    $getPnpDevice = Get-Command -Name "Get-PnpDevice" -ErrorAction SilentlyContinue
    if ($null -ne $getPnpDevice) {
        $ports = @(Get-PnpDevice -Class "Ports" -ErrorAction SilentlyContinue |
            Select-Object Status, Class, FriendlyName, InstanceId, ProblemCode)
        $usb = @(Get-PnpDevice -Class "USB" -ErrorAction SilentlyContinue |
            Where-Object { $_.FriendlyName -match "AsReader|Cognex|DataMan|RFID|Serial|USB" } |
            Select-Object Status, Class, FriendlyName, InstanceId, ProblemCode)
    }

    $serialPorts = @()
    try {
        $serialPorts = @(Get-CimInstance -ClassName Win32_SerialPort -ErrorAction Stop |
            Select-Object DeviceID, Name, Description, Status, PNPDeviceID)
    }
    catch {
        try {
            $serialPorts = @(Get-WmiObject -Class Win32_SerialPort -ErrorAction Stop |
                Select-Object DeviceID, Name, Description, Status, PNPDeviceID)
        }
        catch {
            $serialPorts = @([pscustomobject]@{
                Error = $_.Exception.Message
            })
        }
    }

    return [pscustomobject]@{
        PnpPorts    = $ports
        PnpUsb      = $usb
        SerialPorts = $serialPorts
    }
}

function Get-PathInventory {
    $entries = @()
    if (-not [string]::IsNullOrWhiteSpace($env:Path)) {
        $entries = @($env:Path -split ";" |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique)
    }
    return $entries
}

function Get-RelativeFileInventory {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Root,

        [string[]] $Names = @()
    )

    if (-not (Test-Path -LiteralPath $Root -PathType Container)) {
        return @()
    }

    $found = @()
    foreach ($name in $Names) {
        $found += @(Get-ChildItem -LiteralPath $Root -Filter $name -File -Recurse `
            -ErrorAction SilentlyContinue |
            Select-Object FullName, Length, LastWriteTime)
    }
    return @($found | Sort-Object FullName -Unique)
}
