[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$classId = "{62A782A4-0AC8-4B8E-BC92-E5D66D1360A8}"
$progId = "SmartMove.OutlookAddIn"
$paths = @(
    "Software\Microsoft\Office\Outlook\Addins\$progId",
    "Software\Classes\$progId",
    "Software\Classes\CLSID\$classId"
)

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

foreach ($view in @([Microsoft.Win32.RegistryView]::Registry32, [Microsoft.Win32.RegistryView]::Registry64)) {
    $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::CurrentUser,
        $view)
    try {
        foreach ($path in $paths) {
            $root.DeleteSubKeyTree($path, $false)
        }
        Remove-SmartMoveResiliencyEntries $root
    }
    finally {
        $root.Dispose()
    }
}

Write-Host "SmartMove wurde für den aktuellen Benutzer deinstalliert." -ForegroundColor Green
Write-Host "Starten Sie das klassische Outlook jetzt neu."
