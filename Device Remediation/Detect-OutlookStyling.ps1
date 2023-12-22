<#
.SYNOPSIS
    PowerShell script to detect an existing Outlook `Stationary and Fonts` Styling.

.EXAMPLE
    .\Detect-OutlookStyling.ps1

.DESCRIPTION
    This PowerShell script is deployed as a detection script using Microsoft Intune remediations.

.LINK
    https://github.com/sjackson0109/Intune/blob/main/Device%20Remediation/Detect-OutlookStyling.ps1

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


Try {
    $ThemeFont = Get-ItemProperty -Path $key -Name "ThemeFont" -ErrorAction Stop | Select-Object -ExpandProperty "ThemeFont"
    If ($ThemeFont -eq "Corporate Branding"){
        Write-Output "Compliant"
        Exit 0
    } 
    Write-Warning "Not Compliant"
    Exit 1
} 
Catch {
    Write-Warning "Not Compliant"
    Exit 1
}