<#
.SYNOPSIS
    PowerShell script to remediate Outlook `Stationary and Fonts` Styling following corporate branding.

.EXAMPLE
    .\Remediate-OutlookStyling.ps1

.DESCRIPTION
    This PowerShell script is deployed as a detection script using Microsoft Intune remediations.

.LINK
    https://github.com/sjackson0109/Intune/blob/main/Device%20Remediation/Remediate-OutlookStyling.ps1

.NOTES
    Version:        1.0.1
    Creation Date:  2023-12-22
    Last Updated:   2023-12-22
    Author:         Joey Verlinden / j0eyv
    Contributor:    Simon Jackson / sjackson0109
#>
Function Get-InstalledMSOfficeVersion{
    [CmdletBinding()]

    ## Determine installed MS Office Version
    $OfficeVersionX32        = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue) | Select-Object -ExpandProperty VersionToReport
    $OfficeVersionX64        = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)

    if ( $OfficeVersionX32 -ne $null -and $OfficeVersionX64 -ne $null) {
        $OfficeVersion = "Both x32 version ($OfficeVersionX32) and x64 version ($OfficeVersionX64) installed!"
    } elseif ($OfficeVersionX32 -eq $null -or $OfficeVersionX64 -eq $null) {
        $OfficeVersion = $OfficeVersionX32 + $OfficeVersionX64
    }
    Return $OfficeVersion.Split(".")[0]
}
$ver = Get-InstalledMSOfficeVersion
$key = "HKCU:\SOFTWARE\Microsoft\Office\$ver.0\Common\mailsettings"



# Instructions you should define the formatting staff need inside MS Outlook. Using: Outlook > Options > Mail > Stationary and Fonts. Set both the COMPOSE and REPLY fonts, font-name, font-colour and font-size only.
# Retrieve the HEX ARRAY values, and Split them into a comma-separated array like so:
$originalValueSimple = (Get-ItemProperty -Path $key -Name "TextFontSimple" -ErrorAction Stop | Select-Object -ExpandProperty "TextFontSimple" | ForEach-Object { '{0:X2}' -f $_ }) -join ','
$originalValueComposeComplex = (Get-ItemProperty -Path $key -Name "ComposeFontComplex" -ErrorAction Stop | Select-Object -ExpandProperty "ComposeFontComplex" | ForEach-Object { '{0:X2}' -f $_ }) -join ','
$originalValueReplyComplex = (Get-ItemProperty -Path $key -Name "ReplyFontComplex" -ErrorAction Stop | Select-Object -ExpandProperty "ReplyFontComplex" | ForEach-Object { '{0:X2}' -f $_ }) -join ','
$originalValueTextComplex = (Get-ItemProperty -Path $key -Name "TextFontComplex" -ErrorAction Stop | Select-Object -ExpandProperty "TextFontComplex" | ForEach-Object { '{0:X2}' -f $_ }) -join ','

# Update these values, AFTER agreeing on the style required. Don't forget to either add it all as one string, or include the + ` operation to wrap lines as a single string
# Convertion from STRING to CHAR ARRAY, and then to HEX ARRAY, and then to BYTE ARRAY is possible. This takes a lot of time for a non-programmer, and it's simply easier to manually set/read/update the values here.
# Note: Below includes: Arial, 10pt, Black
$ValueSimple = "3C,00,00,00,1F,00,00,F8,00,00,00,40,DC,00,00,00,00,00,00,00,00,00,00,00,00,22,43,61,6C,69,62,72,69,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00"
$ValueComposeComplex = "3C,68,74,6D,6C,3E,0D,0A,0D,0A,3C,68,65,61,64,3E,0D,0A,3C,73,74,79,6C,65,3E,0D,0A,0D,0A,20,2F,2A,20,53,74,79,6C,65,20,44,65,66,69,6E,69,74,69,6F,6E,73,20,2A,2F,0D,0A,20,73,70,61,6E,2E,50,65,72,73,6F,6E,61,6C,43,6F,6D,70,6F,73,65,53,74,79,6C,65,0D,0A,09,7B,6D," +`
"73,6F,2D,73,74,79,6C,65,2D,6E,61,6D,65,3A,22,50,65,72,73,6F,6E,61,6C,20,43,6F,6D,70,6F,73,65,20,53,74,79,6C,65,22,3B,0D,0A,09,6D,73,6F,2D,73,74,79,6C,65,2D,74,79,70,65,3A,70,65,72,73,6F,6E,61,6C,2D,63,6F,6D,70,6F,73,65,3B,0D,0A,09,6D,73,6F,2D,73,74,79,6C,65," + `
"2D,6E,6F,73,68,6F,77,3A,79,65,73,3B,0D,0A,09,6D,73,6F,2D,73,74,79,6C,65,2D,75,6E,68,69,64,65,3A,6E,6F,3B,0D,0A,09,6D,73,6F,2D,61,6E,73,69,2D,66,6F,6E,74,2D,73,69,7A,65,3A,31,31,2E,30,70,74,3B,0D,0A,09,6D,73,6F,2D,62,69,64,69,2D,66,6F,6E,74,2D,73,69,7A,65,3A," + `
"31,31,2E,30,70,74,3B,0D,0A,09,66,6F,6E,74,2D,66,61,6D,69,6C,79,3A,22,41,72,69,61,6C,22,2C,73,61,6E,73,2D,73,65,72,69,66,3B,0D,0A,09,6D,73,6F,2D,61,73,63,69,69,2D,66,6F,6E,74,2D,66,61,6D,69,6C,79,3A,41,72,69,61,6C,3B,0D,0A,09,6D,73,6F,2D,68,61,6E,73,69,2D,66," + `
"6F,6E,74,2D,66,61,6D,69,6C,79,3A,41,72,69,61,6C,3B,0D,0A,09,6D,73,6F,2D,62,69,64,69,2D,66,6F,6E,74,2D,66,61,6D,69,6C,79,3A,43,61,6C,69,62,72,69,3B,0D,0A,09,63,6F,6C,6F,72,3A,77,69,6E,64,6F,77,74,65,78,74,3B,0D,0A,09,6D,73,6F,2D,74,65,78,74,2D,61,6E,69,6D,61," + `
"74,69,6F,6E,3A,6E,6F,6E,65,3B,0D,0A,09,66,6F,6E,74,2D,77,65,69,67,68,74,3A,6E,6F,72,6D,61,6C,3B,0D,0A,09,66,6F,6E,74,2D,73,74,79,6C,65,3A,6E,6F,72,6D,61,6C,3B,0D,0A,09,74,65,78,74,2D,64,65,63,6F,72,61,74,69,6F,6E,3A,6E,6F,6E,65,3B,0D,0A,09,74,65,78,74,2D,75," + `
"6E,64,65,72,6C,69,6E,65,3A,6E,6F,6E,65,3B,0D,0A,09,74,65,78,74,2D,64,65,63,6F,72,61,74,69,6F,6E,3A,6E,6F,6E,65,3B,0D,0A,09,74,65,78,74,2D,6C,69,6E,65,2D,74,68,72,6F,75,67,68,3A,6E,6F,6E,65,3B,7D,0D,0A,2D,2D,3E,0D,0A,3C,2F,73,74,79,6C,65,3E,0D,0A,3C,2F,68,65," +`
"61,64,3E,0D,0A,0D,0A,3C,2F,68,74,6D,6C,3E,0D,0A"
$ValueReplyComplex = "3C,68,74,6D,6C,3E,0D,0A,0D,0A,3C,68,65,61,64,3E,0D,0A,3C,73,74,79,6C,65,3E,0D,0A,0D,0A,20,2F,2A,20,53,74,79,6C,65,20,44,65,66,69,6E,69,74,69,6F,6E,73,20,2A,2F,0D,0A,20,73,70,61,6E,2E,50,65,72,73,6F,6E,61,6C,52,65,70,6C,79,53,74,79,6C,65,31,0D,0A,09,7B,6D,73," +`
"6F,2D,73,74,79,6C,65,2D,6E,61,6D,65,3A,22,50,65,72,73,6F,6E,61,6C,20,52,65,70,6C,79,20,53,74,79,6C,65,31,22,3B,0D,0A,09,6D,73,6F,2D,73,74,79,6C,65,2D,74,79,70,65,3A,70,65,72,73,6F,6E,61,6C,2D,72,65,70,6C,79,3B,0D,0A,09,6D,73,6F,2D,73,74,79,6C,65,2D,6E,6F,73," +`
"68,6F,77,3A,79,65,73,3B,0D,0A,09,6D,73,6F,2D,73,74,79,6C,65,2D,75,6E,68,69,64,65,3A,6E,6F,3B,0D,0A,09,6D,73,6F,2D,61,6E,73,69,2D,66,6F,6E,74,2D,73,69,7A,65,3A,31,31,2E,30,70,74,3B,0D,0A,09,6D,73,6F,2D,62,69,64,69,2D,66,6F,6E,74,2D,73,69,7A,65,3A,31,31,2E,30," +`
"70,74,3B,0D,0A,09,66,6F,6E,74,2D,66,61,6D,69,6C,79,3A,22,41,72,69,61,6C,22,2C,73,61,6E,73,2D,73,65,72,69,66,3B,0D,0A,09,6D,73,6F,2D,61,73,63,69,69,2D,66,6F,6E,74,2D,66,61,6D,69,6C,79,3A,41,72,69,61,6C,3B,0D,0A,09,6D,73,6F,2D,68,61,6E,73,69,2D,66,6F,6E,74,2D," +`
"66,61,6D,69,6C,79,3A,41,72,69,61,6C,3B,0D,0A,09,6D,73,6F,2D,62,69,64,69,2D,66,6F,6E,74,2D,66,61,6D,69,6C,79,3A,22,54,69,6D,65,73,20,4E,65,77,20,52,6F,6D,61,6E,22,3B,0D,0A,09,6D,73,6F,2D,62,69,64,69,2D,74,68,65,6D,65,2D,66,6F,6E,74,3A,6D,69,6E,6F,72,2D,62,69," +`
"64,69,3B,0D,0A,09,63,6F,6C,6F,72,3A,77,69,6E,64,6F,77,74,65,78,74,3B,7D,0D,0A,2D,2D,3E,0D,0A,3C,2F,73,74,79,6C,65,3E,0D,0A,3C,2F,68,65,61,64,3E,0D,0A,0D,0A,3C,2F,68,74,6D,6C,3E,0D,0A"

$ValueTextComplex = "3C,68,74,6D,6C,3E,0D,0A,0D,0A,3C,68,65,61,64,3E,0D,0A,3C,73,74,79,6C,65,3E,0D,0A,0D,0A,20,2F,2A,20,53,74,79,6C,65,20,44,65,66,69,6E,69,74,69,6F,6E,73,20,2A,2F,0D,0A,20,70,2E,4D,73,6F,50,6C,61,69,6E,54,65,78,74,2C,20,6C,69,2E,4D,73,6F,50,6C,61,69,6E,54,65,78," +`
"74,2C,20,64,69,76,2E,4D,73,6F,50,6C,61,69,6E,54,65,78,74,0D,0A,09,7B,6D,73,6F,2D,73,74,79,6C,65,2D,6E,6F,73,68,6F,77,3A,79,65,73,3B,0D,0A,09,6D,73,6F,2D,73,74,79,6C,65,2D,70,72,69,6F,72,69,74,79,3A,39,39,3B,0D,0A,09,6D,73,6F,2D,73,74,79,6C,65,2D,6C,69,6E,6B," +`
"3A,22,50,6C,61,69,6E,20,54,65,78,74,20,43,68,61,72,22,3B,0D,0A,09,6D,61,72,67,69,6E,3A,30,63,6D,3B,0D,0A,09,6D,73,6F,2D,70,61,67,69,6E,61,74,69,6F,6E,3A,77,69,64,6F,77,2D,6F,72,70,68,61,6E,3B,0D,0A,09,66,6F,6E,74,2D,73,69,7A,65,3A,31,31,2E,30,70,74,3B,0D,0A," +`
"09,6D,73,6F,2D,62,69,64,69,2D,66,6F,6E,74,2D,73,69,7A,65,3A,31,30,2E,35,70,74,3B,0D,0A,09,66,6F,6E,74,2D,66,61,6D,69,6C,79,3A,22,43,61,6C,69,62,72,69,22,2C,73,61,6E,73,2D,73,65,72,69,66,3B,0D,0A,09,6D,73,6F,2D,66,61,72,65,61,73,74,2D,66,6F,6E,74,2D,66,61,6D," +`
"69,6C,79,3A,43,61,6C,69,62,72,69,3B,0D,0A,09,6D,73,6F,2D,66,61,72,65,61,73,74,2D,74,68,65,6D,65,2D,66,6F,6E,74,3A,6D,69,6E,6F,72,2D,6C,61,74,69,6E,3B,0D,0A,09,6D,73,6F,2D,62,69,64,69,2D,66,6F,6E,74,2D,66,61,6D,69,6C,79,3A,22,54,69,6D,65,73,20,4E,65,77,20,52," +`
"6F,6D,61,6E,22,3B,0D,0A,09,6D,73,6F,2D,62,69,64,69,2D,74,68,65,6D,65,2D,66,6F,6E,74,3A,6D,69,6E,6F,72,2D,62,69,64,69,3B,0D,0A,09,6D,73,6F,2D,66,6F,6E,74,2D,6B,65,72,6E,69,6E,67,3A,31,2E,30,70,74,3B,0D,0A,09,6D,73,6F,2D,6C,69,67,61,74,75,72,65,73,3A,73,74,61," +`
"6E,64,61,72,64,63,6F,6E,74,65,78,74,75,61,6C,3B,0D,0A,09,6D,73,6F,2D,66,61,72,65,61,73,74,2D,6C,61,6E,67,75,61,67,65,3A,45,4E,2D,55,53,3B,7D,0D,0A,2D,2D,3E,0D,0A,3C,2F,73,74,79,6C,65,3E,0D,0A,3C,2F,68,65,61,64,3E,0D,0A,0D,0A,3C,2F,68,74,6D,6C,3E,0D,0A"

# DO SOME WORK WITH THE REGISTRY
If(!(Test-Path $key))
{
    New-Item -Path $key -Force | Out-Null
    New-ItemProperty -Path $key -name "NewTheme" -PropertyType String  -value $null
    New-ItemProperty -Path $key -name "ThemeFont" -PropertyType String  -value "Corporate Branded"
    New-ItemProperty -Path $key -Name "ComposeFontSimple" -Value ([byte[]]$($ValueSimple.Split(',') | % { "0x$_"})) -PropertyType Binary -Force
    New-ItemProperty -Path $key -Name "ReplyFontSimple" -Value ([byte[]]$($ValueSimple.Split(',') | % { "0x$_"})) -PropertyType Binary -Force
    New-ItemProperty -Path $key -Name "TextFontSimple" -Value ([byte[]]$($ValueSimple.Split(',') | % { "0x$_"})) -PropertyType Binary -Force
    New-ItemProperty -Path $key -Name "ComposeFontComplex" -Value ([byte[]]$($ValueComposeComplex.Split(',') | % { "0x$_"})) -PropertyType Binary -Force
    New-ItemProperty -Path $key -Name "ReplyFontComplex" -Value ([byte[]]$($ValueReplyComplex.Split(',') | % { "0x$_"})) -PropertyType Binary -Force
    New-ItemProperty -Path $key -Name "TextFontComplex" -Value ([byte[]]$($ValueTextComplex.Split(',') | % { "0x$_"})) -PropertyType Binary -Force
} Else {
    Set-ItemProperty -Path $key -name "NewTheme" -value $null
    Set-ItemProperty -Path $key -name "ThemeFont" -value "Corporate Branded"
    Set-ItemProperty -Path $key -Name "ComposeFontSimple" -Value [byte[]]$($ValueSimple.Split(',') | % { "0x$_"}) -Force
    Set-ItemProperty -Path $key -Name "ReplyFontSimple" -Value [byte[]]$($ValueSimple.Split(',') | % { "0x$_"}) -Force
    Set-ItemProperty -Path $key -Name "TextFontSimple" -Value [byte[]]$($ValueSimple.Split(',') | % { "0x$_"}) -Force
    Set-ItemProperty -Path $key -Name "ComposeFontComplex" -Value ([byte[]]$($ValueComposeComplex.Split(',') | % { "0x$_"})) -Force
    Set-ItemProperty -Path $key -Name "ReplyFontComplex" -Value ([byte[]]$($ValueReplyComplex.Split(',') | % { "0x$_"})) -Force
    Set-ItemProperty -Path $key -Name "TextFontComplex" -Value ([byte[]]$($ValueTextComplex.Split(',') | % { "0x$_"})) -Force
}