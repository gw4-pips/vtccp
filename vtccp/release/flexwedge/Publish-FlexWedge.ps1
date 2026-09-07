param([string]$Configuration = "Release")
$ErrorActionPreference = "Stop"
$root = Resolve-Path (Join-Path $PSScriptRoot "..\..")
$sdk = Join-Path $root "lib\asreader-p3xu-sdk-1.3.0\AsReaderP3xU.dll"
if (!(Test-Path $sdk)) { throw "Required vendor SDK DLL is absent: $sdk. Obtain redistribution permission and the DLL from AsReader before packaging." }
$project = Join-Path $root "FlexWedgeApp\FlexWedgeApp.csproj"
$tests = Join-Path $root "DeviceInterface.Tests\DeviceInterface.Tests.csproj"
$out = Join-Path $root "artifacts\flexwedge\publish"
$zip = Join-Path $root "artifacts\flexwedge\FlexWedge-1.0.0-win-x64.zip"
dotnet restore $project
dotnet restore $tests
dotnet test $tests --no-restore --filter "FullyQualifiedName~FlexWedge"
Remove-Item $out -Recurse -Force -ErrorAction SilentlyContinue
dotnet publish $project -c $Configuration -r win-x64 --self-contained true -p:PublishSingleFile=false -o $out
Copy-Item $sdk $out -Force
Copy-Item (Join-Path $PSScriptRoot "install.ps1") $out -Force
Copy-Item (Join-Path $PSScriptRoot "uninstall.ps1") $out -Force
Copy-Item (Join-Path $root "references\flexwedge\*.md") $out -Force
Remove-Item $zip -Force -ErrorAction SilentlyContinue
Compress-Archive -Path (Join-Path $out "*") -DestinationPath $zip
Write-Host "Created $zip"