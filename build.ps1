[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"
$project = Join-Path $PSScriptRoot "SmartMove.OutlookAddIn\SmartMove.OutlookAddIn.csproj"

dotnet build $project --configuration $Configuration
if ($LASTEXITCODE -ne 0) {
    throw "Der Build ist fehlgeschlagen."
}

$dll = Join-Path $PSScriptRoot "SmartMove.OutlookAddIn\bin\$Configuration\net48\SmartMove.OutlookAddIn.dll"
Write-Host "SmartMove wurde erstellt: $dll" -ForegroundColor Green
