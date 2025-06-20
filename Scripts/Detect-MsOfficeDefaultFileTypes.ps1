# Detection Script for Office default save formats (.xlsx, .docx, .pptx)

function Get-OfficeVersion {
    $officeRoot = "HKCU:\Software\Microsoft\Office"
    $versions = Get-ChildItem -Path $officeRoot -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '\\\d+(\.\d+)?$' } |
        ForEach-Object { [version]($_.PSChildName) } |
        Sort-Object -Descending

    return $versions | Where-Object { $_ -ge [version]"14.0" } | Select-Object -First 1
}

$officeVer = Get-OfficeVersion
if (-not $officeVer) {
    Write-Output "Office not installed"
    exit 0  # Consider compliant if Office is not present
}

$verString = $officeVer.ToString()

$apps = @{
    "Excel"      = @{ Path = "HKCU:\Software\Microsoft\Office\$verString\Excel\Options";      Property = "DefaultFormat"; Expected = 51 }
    "Word"       = @{ Path = "HKCU:\Software\Microsoft\Office\$verString\Word\Options";       Property = "DefaultFormat"; Expected = 16 }
    "PowerPoint" = @{ Path = "HKCU:\Software\Microsoft\Office\$verString\PowerPoint\Options"; Property = "DefaultFormat"; Expected = 24 }
}

$nonCompliant = @()

foreach ($app in $apps.Keys) {
    $info = $apps[$app]
    $actual = ""

    if (Test-Path $info.Path) {
        $actual = (Get-ItemProperty -Path $info.Path -Name $info.Property -ErrorAction SilentlyContinue).$($info.Property)
    }

    if ($actual -ne $info.Expected) {
        $nonCompliant += $app
    }
}

if ($nonCompliant.Count -eq 0) {
    Write-Output "Compliant"
    exit 0
} else {
    Write-Output "Non-compliant: " + ($nonCompliant -join ", ")
    exit 1
}
