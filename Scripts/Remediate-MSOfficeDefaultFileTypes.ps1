# Remediation Script for Office default save formats (.xlsx, .docx, .pptx)

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
    exit 0
}

$verString = $officeVer.ToString()

$apps = @{
    "Excel"      = @{ Path = "HKCU:\Software\Microsoft\Office\$verString\Excel\Options";      Property = "DefaultFormat"; Value = 51 }
    "Word"       = @{ Path = "HKCU:\Software\Microsoft\Office\$verString\Word\Options";       Property = "DefaultFormat"; Value = 16 }
    "PowerPoint" = @{ Path = "HKCU:\Software\Microsoft\Office\$verString\PowerPoint\Options"; Property = "DefaultFormat"; Value = 24 }
}

foreach ($app in $apps.Keys) {
    $info = $apps[$app]

    if (-not (Test-Path $info.Path)) {
        New-Item -Path $info.Path -Force | Out-Null
    }

    Set-ItemProperty -Path $info.Path -Name $info.Property -Value $info.Value -Force
}
