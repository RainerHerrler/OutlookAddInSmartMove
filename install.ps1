[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"
$classId = "{62A782A4-0AC8-4B8E-BC92-E5D66D1360A8}"
$progId = "SmartMove.OutlookAddIn"
$className = "SmartMove.OutlookAddIn.SmartMoveAddIn"
$assemblyName = "SmartMove.OutlookAddIn, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
$runtimeVersion = "v4.0.30319"
$version = "1.0.0.0"

if (-not $SkipBuild) {
    & (Join-Path $PSScriptRoot "build.ps1") -Configuration $Configuration
}

$dll = Join-Path $PSScriptRoot "SmartMove.OutlookAddIn\bin\$Configuration\net48\SmartMove.OutlookAddIn.dll"
if (-not (Test-Path $dll)) {
    throw "Add-in-DLL nicht gefunden: $dll"
}

$codeBase = ([System.Uri]$dll).AbsoluteUri
function Set-ComValues([Microsoft.Win32.RegistryKey]$Key) {
    $Key.SetValue("", "mscoree.dll", [Microsoft.Win32.RegistryValueKind]::String)
    $Key.SetValue("ThreadingModel", "Both", [Microsoft.Win32.RegistryValueKind]::String)
    $Key.SetValue("Class", $className, [Microsoft.Win32.RegistryValueKind]::String)
    $Key.SetValue("Assembly", $assemblyName, [Microsoft.Win32.RegistryValueKind]::String)
    $Key.SetValue("RuntimeVersion", $runtimeVersion, [Microsoft.Win32.RegistryValueKind]::String)
    $Key.SetValue("CodeBase", $codeBase, [Microsoft.Win32.RegistryValueKind]::String)
}

function Remove-SmartMoveResiliencyEntries([Microsoft.Win32.RegistryKey]$Root) {
    foreach ($subPath in @(
        "Software\Microsoft\Office\16.0\Outlook\Resiliency\CrashingAddinList",
        "Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems")) {
        $key = $Root.OpenSubKey($subPath, $true)
        if ($null -eq $key) {
            continue
        }

        try {
            foreach ($name in @($key.GetValueNames())) {
                $value = $key.GetValue($name)
                if ($value -is [byte[]]) {
                    $text = [Text.Encoding]::Unicode.GetString($value)
                    if ($text.IndexOf("smartmove.outlookaddin", [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                        $key.DeleteValue($name, $false)
                    }
                }
            }
        }
        finally {
            $key.Dispose()
        }
    }
}

# Register both views so the add-in works with 32- and 64-bit classic Outlook.
foreach ($view in @([Microsoft.Win32.RegistryView]::Registry32, [Microsoft.Win32.RegistryView]::Registry64)) {
    $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::CurrentUser,
        $view)
    try {
        Remove-SmartMoveResiliencyEntries $root

        $clsidKey = $root.CreateSubKey("Software\Classes\CLSID\$classId")
        $clsidKey.SetValue("", $className, [Microsoft.Win32.RegistryValueKind]::String)
        $clsidKey.Dispose()

        $inprocKey = $root.CreateSubKey("Software\Classes\CLSID\$classId\InprocServer32")
        Set-ComValues $inprocKey
        $inprocKey.Dispose()

        $versionKey = $root.CreateSubKey("Software\Classes\CLSID\$classId\InprocServer32\$version")
        Set-ComValues $versionKey
        $versionKey.Dispose()

        $clsidProgIdKey = $root.CreateSubKey("Software\Classes\CLSID\$classId\ProgId")
        $clsidProgIdKey.SetValue("", $progId, [Microsoft.Win32.RegistryValueKind]::String)
        $clsidProgIdKey.Dispose()

        $progIdKey = $root.CreateSubKey("Software\Classes\$progId")
        $progIdKey.SetValue("", $className, [Microsoft.Win32.RegistryValueKind]::String)
        $progIdKey.Dispose()

        $progIdClsidKey = $root.CreateSubKey("Software\Classes\$progId\CLSID")
        $progIdClsidKey.SetValue("", $classId, [Microsoft.Win32.RegistryValueKind]::String)
        $progIdClsidKey.Dispose()

        $addInKey = $root.CreateSubKey("Software\Microsoft\Office\Outlook\Addins\$progId")
        $addInKey.SetValue("FriendlyName", "SmartMove", [Microsoft.Win32.RegistryValueKind]::String)
        $addInKey.SetValue("Description", "Schlägt passende Zielordner für bearbeitete E-Mails vor.", [Microsoft.Win32.RegistryValueKind]::String)
        $addInKey.SetValue("LoadBehavior", 3, [Microsoft.Win32.RegistryValueKind]::DWord)
        $addInKey.Dispose()
    }
    finally {
        $root.Dispose()
    }
}

Write-Host "SmartMove wurde für den aktuellen Benutzer installiert." -ForegroundColor Green
Write-Host "Starten Sie das klassische Outlook jetzt neu."
