<#
.SYNOPSIS
    PowerShell script to download and deploy the lates published outlook signautres, off github.

.EXAMPLE
    .\Detect-OutlookSignatures.ps1

.DESCRIPTION
    This PowerShell script is deployed as a detection script using Microsoft Intune remediations.

.LINK
    https://github.com/sjackson0109/EmailTemplates/blob/main/Scripts/Detect-OutlookSignatures.ps1

.LINK
    https://github.com/Set-OutlookSignatures/Set-OutlookSignatures
    
.LINK
    https://learn.microsoft.com/en-us/microsoft-365/security/defender-endpoint/overview-endpoint-detection-response

.NOTES
    Version:        1.0.7
    Creation Date:  2023-11-07
    Last Updated:   2024-01-04
    Author:         Simon Jackson / sjackson0109
    Contact:        simon@jacksonfamily.me
#>
#Look in the localappdata\temp folder
$tempDir = "$($env:TEMP)\OutlookSignatures"
$logFile = "$tempDir\Set-OutlookSignatures.log"
$addHours = 0 # Start with 1 hours, after a couple of weeks move it to 24 hours

# Check if the log file exists
If (Test-Path $logFile ){
    If ( $(Get-Item $logFile).LastWriteTime -gt $(Get-Date).AddHours(-$addHours) ) {
        Write-Host "Log found with timestamp $($(Get-Item $logFile).LastWriteTime)"
        Write-Host "This is old"
        Write-Output "Compliant"
        exit 0
    } Else {
        Write-Host "Log found with timestamp $($(Get-Item $logFile).LastWriteTime)"
        Write-Host "This is recent"
        Write-Output "Non Compliant"
        exit 1
    }
} Else {
    Write-Warning "NO log file found"
    Write-Output "Non Compliant"
    exit 1
}