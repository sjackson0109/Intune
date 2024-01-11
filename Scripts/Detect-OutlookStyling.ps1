<#
.SYNOPSIS
    PowerShell script to detect an existing Outlook `Stationary and Fonts` Styling.

.EXAMPLE
    .\Detect-OutlookStyling.ps1

.DESCRIPTION
    This PowerShell script is deployed as a detection script using Microsoft Intune remediations.

.LINK
    https://github.com/sjackson0109/Intune/blob/main/Scripts/Detect-OutlookStyling.ps1

.NOTES
    Version:        1.0.2
    Creation Date:  2022-12-11
    Last Updated:   2024-01-05
    Inspiration:    Joey Verlinden / j0eyv
    Author:         Simon Jackson / sjackson0109
#>
# Apply your own Corporate Branding
# BrandedFontFamily: "Arial", "Tahoma", "Calibri" etc
# BrandedFontSize: "8.0", "8.5", "9.0", "9.5" ....
# BrandedFontColor: "#fff", "rgb(243,243,243)"", "black", "red"
$BrandingFontFamily = "Arial" 
$BrandedFontSize = "10.0"
$BrandedFontColor = "Black"

# Function to get the current installation version of ms office
function Get-InstalledMSOfficeVersion{
    [CmdletBinding()]
    ## Determine installed MS Office Version
    $OfficeVersionX32        = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue) | Select-Object -ExpandProperty VersionToReport
    $OfficeVersionX64        = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)
    if ( $OfficeVersionX32 -ne $null -and $OfficeVersionX64 -ne $null) {
        $OfficeVersion = "Both x32 version ($OfficeVersionX32) and x64 version ($OfficeVersionX64) installed!"
    } elseif ($OfficeVersionX32 -eq $null -or $OfficeVersionX64 -eq $null) {
        $OfficeVersion = $OfficeVersionX32 + $OfficeVersionX64
    }
    return $OfficeVersion.Split(".")[0]
}

# Function to get a registry value
function Get-BinaryRegistryValue {
    param ( [string]$key, [string]$name )
    try {
        $registryValue = try { Get-ItemProperty -Path $key -Name $name -errorAction SilentlyContinue } catch { "00" }
        $binaryData = $registryValue.$name -split '\s' | ForEach-Object { [byte]($_ -replace '0x','') }
        return $binaryData
    }
    catch { Write-Error "Error getting registry value: $_" }
}

# Function to set a registry value with binary data
function BinaryRegistryValue {
    param ( [string]$Key, [string]$Name, [byte[]]$Value )
    try {
        # Create registry key if not exists
        if (-not (Test-Path $Key)) { New-Item -Path $Key -Force }
        # Use Get/Set-ItemProperty to set the registry value
        New-ItemProperty -Path $Key -Name $Name -PropertyType Binary -Value $Value -Force | Out-Null
        Write-Host "Registry value set successfully."
    } catch { Write-Error "Error setting registry value: $_" }
}

# Function to convert binary data to UTF-8 string
function Convert-BinaryToUTF8String {
    param ( [byte[]]$binaryData )
    try {
        if ($binaryData -eq $null) {
            Write-Error "Binary data is required."
            return
        }
        $utf8String = [System.Text.Encoding]::UTF8.GetString($binaryData)
        return $utf8String
    }
    catch { Write-Error "Error converting binary data to UTF-8 string: $_" }
}

# Function to convert hexadecimal string to UTF-8 string
function Convert-HexToUTF8String {
    param([string]$hexString)
    try {
        $hexPairs = $hexString -replace ',', '' -split '(..)' | Where-Object { $_ }
        $byteArray = [byte[]]@($hexPairs | ForEach-Object { [byte]([Convert]::ToByte($_, 16)) })
        $utf8String = [System.Text.Encoding]::UTF8.GetString($byteArray)
        return $utf8String
    }
    catch { Write-Error "Error converting hexadecimal to UTF-8 string: $_" }
}

# Function to Get the Outlook Style (Default/Current) and return HTML/UTF8 values
function Get-OutlookStyle {
    param( [string]$Name, [string]$Selection )
    # Get Office Version, and Specify MailSettings registry location
    $ver = Get-InstalledMSOfficeVersion
    $registryPath = "HKCU:\SOFTWARE\Microsoft\Office\$ver.0\Common\mailsettings"
    $defaultValues = @{
        ComposeFontComplex = "3c,68,74,6d,6c,3e,0d,0a,0d,0a,3c,68,65,61,64,3e,0d,0a," + `
            "3c,73,74,79,6c,65,3e,0d,0a,0d,0a,20,2f,2a,20,53,74,79,6c,65,20,44,65,66,69," + `
            "6e,69,74,69,6f,6e,73,20,2a,2f,0d,0a,20,73,70,61,6e,2e,50,65,72,73,6f,6e,61," + `
            "6c,43,6f,6d,70,6f,73,65,53,74,79,6c,65,0d,0a,09,7b,6d,73,6f,2d,73,74,79,6c," + `
            "65,2d,6e,61,6d,65,3a,22,50,65,72,73,6f,6e,61,6c,20,43,6f,6d,70,6f,73,65,20," + `
            "53,74,79,6c,65,22,3b,0d,0a,09,6d,73,6f,2d,73,74,79,6c,65,2d,74,79,70,65,3a," + `
            "70,65,72,73,6f,6e,61,6c,2d,63,6f,6d,70,6f,73,65,3b,0d,0a,09,6d,73,6f,2d,73," + `
            "74,79,6c,65,2d,6e,6f,73,68,6f,77,3a,79,65,73,3b,0d,0a,09,6d,73,6f,2d,73,74," + `
            "79,6c,65,2d,75,6e,68,69,64,65,3a,6e,6f,3b,0d,0a,09,6d,73,6f,2d,61,6e,73,69," + `
            "2d,66,6f,6e,74,2d,73,69,7a,65,3a,31,31,2e,30,70,74,3b,0d,0a,09,6d,73,6f,2d," + `
            "62,69,64,69,2d,66,6f,6e,74,2d,73,69,7a,65,3a,31,31,2e,30,70,74,3b,0d,0a,09," + `
            "66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,22,43,61,6c,69,62,72,69,22,2c,73,61,6e," + `
            "73,2d,73,65,72,69,66,3b,0d,0a,09,6d,73,6f,2d,61,73,63,69,69,2d,66,6f,6e,74," + `
            "2d,66,61,6d,69,6c,79,3a,43,61,6c,69,62,72,69,3b,0d,0a,09,6d,73,6f,2d,61,73," + `
            "63,69,69,2d,74,68,65,6d,65,2d,66,6f,6e,74,3a,6d,69,6e,6f,72,2d,6c,61,74,69," + `
            "6e,3b,0d,0a,09,6d,73,6f,2d,68,61,6e,73,69,2d,66,6f,6e,74,2d,66,61,6d,69,6c," + `
            "79,3a,43,61,6c,69,62,72,69,3b,0d,0a,09,6d,73,6f,2d,68,61,6e,73,69,2d,74,68," + `
            "65,6d,65,2d,66,6f,6e,74,3a,6d,69,6e,6f,72,2d,6c,61,74,69,6e,3b,0d,0a,09,6d," + `
            "73,6f,2d,62,69,64,69,2d,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,22,54,69,6d,65," + `
            "73,20,4e,65,77,20,52,6f,6d,61,6e,22,3b,0d,0a,09,6d,73,6f,2d,62,69,64,69,2d," + `
            "74,68,65,6d,65,2d,66,6f,6e,74,3a,6d,69,6e,6f,72,2d,62,69,64,69,3b,0d,0a,09," + `
            "63,6f,6c,6f,72,3a,77,69,6e,64,6f,77,74,65,78,74,3b,7d,0d,0a,2d,2d,3e,0d,0a," + `
            "3c,2f,73,74,79,6c,65,3e,0d,0a,3c,2f,68,65,61,64,3e,0d,0a,0d,0a,3c,2f,68,74," + `
            "6d,6c,3e,0d,0a"
        ComposeFontSimple = "3c,00,00,00,1f,00,00,f8,00,00,00,40,dc,00,00,00,00,00," + `
            "00,00,00,00,00,00,00,22,43,61,6c,69,62,72,69,00,00,00,00,00,00,00,00,00,00," + `
            "00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00"
        ReplyFontComplex = "3c,68,74,6d,6c,3e,0d,0a,0d,0a,3c,68,65,61,64,3e,0d,0a," + `
            "3c,73,74,79,6c,65,3e,0d,0a,0d,0a,20,2f,2a,20,53,74,79,6c,65,20,44,65,66,69," + `
            "6e,69,74,69,6f,6e,73,20,2a,2f,0d,0a,20,73,70,61,6e,2e,50,65,72,73,6f,6e,61," + `
            "6c,52,65,70,6c,79,53,74,79,6c,65,0d,0a,09,7b,6d,73,6f,2d,73,74,79,6c,65,2d," + `
            "6e,61,6d,65,3a,22,50,65,72,73,6f,6e,61,6c,20,52,65,70,6c,79,20,53,74,79,6c," + `
            "65,22,3b,0d,0a,09,6d,73,6f,2d,73,74,79,6c,65,2d,74,79,70,65,3a,70,65,72,73," + `
            "6f,6e,61,6c,2d,72,65,70,6c,79,3b,0d,0a,09,6d,73,6f,2d,73,74,79,6c,65,2d,6e," + `
            "6f,73,68,6f,77,3a,79,65,73,3b,0d,0a,09,6d,73,6f,2d,73,74,79,6c,65,2d,75,6e," + `
            "68,69,64,65,3a,6e,6f,3b,0d,0a,09,6d,73,6f,2d,61,6e,73,69,2d,66,6f,6e,74,2d," + `
            "73,69,7a,65,3a,31,31,2e,30,70,74,3b,0d,0a,09,6d,73,6f,2d,62,69,64,69,2d,66," + `
            "6f,6e,74,2d,73,69,7a,65,3a,31,31,2e,30,70,74,3b,0d,0a,09,66,6f,6e,74,2d,66," + `
            "61,6d,69,6c,79,3a,22,43,61,6c,69,62,72,69,22,2c,73,61,6e,73,2d,73,65,72,69," + `
            "66,3b,0d,0a,09,6d,73,6f,2d,61,73,63,69,69,2d,66,6f,6e,74,2d,66,61,6d,69,6c," + `
            "79,3a,43,61,6c,69,62,72,69,3b,0d,0a,09,6d,73,6f,2d,61,73,63,69,69,2d,74,68," + `
            "65,6d,65,2d,66,6f,6e,74,3a,6d,69,6e,6f,72,2d,6c,61,74,69,6e,3b,0d,0a,09,6d," + `
            "73,6f,2d,68,61,6e,73,69,2d,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,43,61,6c,69," + `
            "62,72,69,3b,0d,0a,09,6d,73,6f,2d,68,61,6e,73,69,2d,74,68,65,6d,65,2d,66,6f," + `
            "6e,74,3a,6d,69,6e,6f,72,2d,6c,61,74,69,6e,3b,0d,0a,09,6d,73,6f,2d,62,69,64," + `
            "69,2d,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,22,54,69,6d,65,73,20,4e,65,77,20," + `
            "52,6f,6d,61,6e,22,3b,0d,0a,09,6d,73,6f,2d,62,69,64,69,2d,74,68,65,6d,65,2d," + `
            "66,6f,6e,74,3a,6d,69,6e,6f,72,2d,62,69,64,69,3b,0d,0a,09,63,6f,6c,6f,72,3a," + `
            "77,69,6e,64,6f,77,74,65,78,74,3b,7d,0d,0a,2d,2d,3e,0d,0a,3c,2f,73,74,79,6c," + `
            "65,3e,0d,0a,3c,2f,68,65,61,64,3e,0d,0a,0d,0a,3c,2f,68,74,6d,6c,3e,0d,0a"
        ReplyFontSimple = "3c,00,00,00,1f,00,00,f8,00,00,00,40,dc,00,00,00,00,00,00," + `
            "00,00,00,00,00,00,22,43,61,6c,69,62,72,69,00,00,00,00,00,00,00,00,00,00,00," + `
            "00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00"
        TextFontComplex = "3c,68,74,6d,6c,3e,0d,0a,0d,0a,3c,68,65,61,64,3e,0d,0a,3c," + `
            "73,74,79,6c,65,3e,0d,0a,0d,0a,20,2f,2a,20,53,74,79,6c,65,20,44,65,66,69,6e," + `
            "69,74,69,6f,6e,73,20,2a,2f,0d,0a,20,70,2e,4d,73,6f,50,6c,61,69,6e,54,65,78," + `
            "74,2c,20,6c,69,2e,4d,73,6f,50,6c,61,69,6e,54,65,78,74,2c,20,64,69,76,2e,4d," + `
            "73,6f,50,6c,61,69,6e,54,65,78,74,0d,0a,09,7b,6d,73,6f,2d,73,74,79,6c,65,2d," + `
            "6e,6f,73,68,6f,77,3a,79,65,73,3b,0d,0a,09,6d,73,6f,2d,73,74,79,6c,65,2d,70," + `
            "72,69,6f,72,69,74,79,3a,39,39,3b,0d,0a,09,6d,73,6f,2d,73,74,79,6c,65,2d,6c," + `
            "69,6e,6b,3a,22,50,6c,61,69,6e,20,54,65,78,74,20,43,68,61,72,22,3b,0d,0a,09," + `
            "6d,61,72,67,69,6e,3a,30,63,6d,3b,0d,0a,09,6d,73,6f,2d,70,61,67,69,6e,61,74," + `
            "69,6f,6e,3a,77,69,64,6f,77,2d,6f,72,70,68,61,6e,3b,0d,0a,09,66,6f,6e,74,2d," + `
            "73,69,7a,65,3a,31,31,2e,30,70,74,3b,0d,0a,09,6d,73,6f,2d,62,69,64,69,2d,66," + `
            "6f,6e,74,2d,73,69,7a,65,3a,31,30,2e,35,70,74,3b,0d,0a,09,66,6f,6e,74,2d,66," + `
            "61,6d,69,6c,79,3a,22,43,61,6c,69,62,72,69,22,2c,73,61,6e,73,2d,73,65,72,69," + `
            "66,3b,0d,0a,09,6d,73,6f,2d,66,61,72,65,61,73,74,2d,66,6f,6e,74,2d,66,61,6d," + `
            "69,6c,79,3a,43,61,6c,69,62,72,69,3b,0d,0a,09,6d,73,6f,2d,66,61,72,65,61,73," + `
            "74,2d,74,68,65,6d,65,2d,66,6f,6e,74,3a,6d,69,6e,6f,72,2d,6c,61,74,69,6e,3b," + `
            "0d,0a,09,6d,73,6f,2d,62,69,64,69,2d,66,6f,6e,74,2d,66,61,6d,69,6c,79,3a,22," + `
            "54,69,6d,65,73,20,4e,65,77,20,52,6f,6d,61,6e,22,3b,0d,0a,09,6d,73,6f,2d,62," + `
            "69,64,69,2d,74,68,65,6d,65,2d,66,6f,6e,74,3a,6d,69,6e,6f,72,2d,62,69,64,69," + `
            "3b,0d,0a,09,6d,73,6f,2d,66,6f,6e,74,2d,6b,65,72,6e,69,6e,67,3a,31,2e,30,70," + `
            "74,3b,0d,0a,09,6d,73,6f,2d,6c,69,67,61,74,75,72,65,73,3a,73,74,61,6e,64,61," + `
            "72,64,63,6f,6e,74,65,78,74,75,61,6c,3b,0d,0a,09,6d,73,6f,2d,66,61,72,65,61," + `
            "73,74,2d,6c,61,6e,67,75,61,67,65,3a,45,4e,2d,55,53,3b,7d,0d,0a,2d,2d,3e,0d," + `
            "0a,3c,2f,73,74,79,6c,65,3e,0d,0a,3c,2f,68,65,61,64,3e,0d,0a,0d,0a,3c,2f,68," + `
            "74,6d,6c,3e,0d,0a"
        TextFontSimple = "3c,00,00,00,1f,00,00,f8,00,00,00,40,dc,00,00,00,00,00,00," + `
            "00,00,00,00,00,00,22,43,61,6c,69,62,72,69,00,00,00,00,00,00,00,00,00,00,00," + `
            "00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00"
    }

    if ($Selection -eq "Current" -or $Selection -eq "current") {
        $binaryData = Get-BinaryRegistryValue -Key $registryPath -Name $Name
        $hexString = -join ($binaryData | ForEach-Object { "{0:X2}" -f $_ })
    } else { $hexString = $defaultValues[$Name] }
    # Convert the HEX to UTF8 before returning.
    return $( Convert-HexToUTF8String -hexString $hexString )
}

# Define the list of style names we are looping over..
$styleNames = @( "ComposeFontSimple", "ComposeFontComplex", "ReplyFontSimple", "ReplyFontComplex", "TextFontSimple", "TextFontComplex" )

$counter = $styleNames.length

# Iterate over each style name, evaluate if the style has changed or not, if not update it.
ForEach ($styleName in $styleNames) {
    $currentStyle = Get-OutlookStyle -Name $styleName -Selection "Current"
    $defaultStyle = Get-OutlookStyle -Name $styleName
    # Check if the current style is equal to the default style
    if ($currentStyle -eq $defaultStyle) {
        Write-Output "$styleName is factory default. Remediation required."
        $counter--  # Decrease the counter when a style is compliant
    } else {
        Write-Output "$styleName style has been customised already. Exiting."
    }
}

# Validation
# Counter value should not change (6), if the branding is compliant.
# Unless all 6x StyleNames are compliant, the entire script should state non-compliant
if ( $counter -eq 6 ) {
    Write-Output "Compliant"
    Exit 0
} else {
    Write-Warning "Not Compliant"
    Exit 1
}