<#
.SYNOPSIS
    PowerShell script to detect an existing Email Signatures log, from Set-OutlookSignatures script.

.EXAMPLE
    .\Detect-EmailSignatures.ps1

.DESCRIPTION
    This PowerShell script is deployed as a detection script using Microsoft Intune remediations.

.LINK
    https://github.com/sjackson0109/Intune/blob/main/Device%20Remediation/Detect-EmailSignatures.ps1

.LINK
    https://github.com/Set-OutlookSignatures/Set-OutlookSignatures
    
.LINK
    https://learn.microsoft.com/en-us/microsoft-365/security/defender-endpoint/overview-endpoint-detection-response

.NOTES
    Version:        1.0.3
    Creation Date:  2023-11-07
    Last Updated:   2023-12-22
    Author:         Simon Jackson / sjackson0109
#>
#Look in the localappdata\temp folder
$temp = $(Get-Location).path
$logFile = "$temp\Set-OutlookSignatures.log"
$addHours = 0 # Start with 2 hours, after a couple of weeks move it to 24 hours

# Check if the log file exists
If (Test-Path $logFile ){
    If ( $(Get-Item $logFile).LastWriteTime -gt $(Get-Date).AddHours(-$addHours) ) {
        Write-Host "NEW log found"
        Write-Output "Compliant"
        exit 0
    } Else {
        Write-Warning "Not Compliant: OLD log file"
        exit 1
    }
}
Else {
    Write-Warning "Not Compliant: no log file found"
    exit 1
}