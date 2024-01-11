<#
.SYNOPSIS
    PowerShell script to remediate an pre/non existing Email Signatures; from Set-OutlookSignatures script.

.EXAMPLE
    .\Remediate-OutlookSignatures.ps1

.DESCRIPTION
    This PowerShell script is deployed as a remediation script using Microsoft Intune remediations.

.LINK
    https://github.com/sjackson0109/EmailTemplates/blob/main/Scripts/Remediate-OutlookSignatures.ps1

.LINK
    https://github.com/Set-OutlookSignatures/Set-OutlookSignatures

.LINK
    https://learn.microsoft.com/en-us/mem/intune/fundamentals/remediations

.NOTES
    Version:        1.0.7
    Creation Date:  2023-11-07
    Last Updated:   2024-01-04
    Author:         Simon Jackson (sjackson0109) 
#>
#$tempDir = $(Get-Location).path
$tempDir = "$($env:TEMP)\OutlookSignatures"
New-Item -ItemType Directory -Force -Path $tempDir -ErrorAction SilentlyContinue
$logFile = "$tempDir\Set-OutlookSignatures.log"
Start-Transcript $logFile -Force

# Variables for Download and Extract
$githubProductOrg = "Set-OutlookSignatures"
$githubProductRepo = "Set-OutlookSignatures"
$githubTemplateOrg = "sjackson0109"         # update this to your own org 
$githubTemplateRepo = "EmailTemplates"      # update this to your own repo

# Product Variables (standard)
$graphOnly = "true"
$SetOofMsg = "false"
$CreateRtfSignatures = "true"
$CreateTxtSignatures = "true"
$SignaturesForAutomappedAndAdditionalMailboxes = "true"

# Product Variables (premium, req benefactor circle)
$DocxHighResImageConversion = "false"
$SetCurrentUserOutlookWebSignature = "true"
$MirrorLocalSignaturesToCloud = "true"
$DeleteUserCreatedSignatures = "true"  #REQ TRUE FOR GO-LIVE
$DeleteScriptCreatedSignaturesWithoutTemplate = "true"


# Init
# Obtain the latest release off each github project  -- note: latest is always array item 0
$productRelease = Invoke-WebRequest -Uri "https://api.github.com/repos/$githubProductOrg/$githubProductRepo/releases/latest" -UseBasicParsing | ConvertFrom-Json
$productUrl = $productRelease.assets.browser_download_url
$productPublished = $productRelease.published_at
$productVersion = $productRelease.tag_name

$templateRelease = Invoke-WebRequest -Uri "https://api.github.com/repos/$githubTemplateOrg/$githubTemplateRepo/releases/latest" -UseBasicParsing | ConvertFrom-Json
$templateUrl = $templateRelease.zipball_url
$templatePublished = $templateRelease.published_at
$templateVersion = $templateRelease.tag_name

# Specify the file-system of the downloaded targets
$productRelease | Out-File "$tempDir\$githubProductRepo.json"
$productZip = "$tempDir\$githubProductRepo-$productVersion.zip"
$productPath = "$tempDir\$githubProductRepo-$productVersion" -replace '-v' , '_v'

$templateRelease | Out-File "$tempDir\$githubTemplateRepo.json"
$templateZip = "$tempDir\$githubTemplateRepo-$templateVersion.zip"
$templatePath = "$tempDir\$githubTemplateRepo-$templateVersion" 

Add-Type -AssemblyName System.IO.Compression.FileSystem

# Check if the latest version is already downloaded, clean up the file-system and download+extract, or just extract again
If (Test-Path "$productPath"){
    Write-Host "Deleting previously extracted files inside $productPath"
    Write-Output "Cleaned up ProductPath"
    Remove-Item $productPath -recurse -Force
} else {
    Write-Host "Downloading $productUrl to $productZip"
    Invoke-WebRequest $productUrl -Out "$productZip"
}

If (Test-Path "$templatePath"){
    Write-Host "Deleting previously extracted files inside $templatePath"
    Write-Output "Cleaned up TemplatePath"
    Remove-Item $templatePath -recurse -Force
} else {
    Write-Host "Downloading $templateUrl to $templateZip"
    Invoke-WebRequest $templateUrl -Out "$templateZip"
}

# A fresh Extraction of the zipball files to the temp directory, filename encoding needs converting to ascii, not utf8.
# Note: some errors with file-name length when testing with my user docs area. C:\WINDOWS\IMECache\HealthScripts\(GUID)\ is just as long, so skip errors. Only signature samples anyway, don't need them.
Write-Host "Extracting $productZip"
try { [System.IO.Compression.ZipFile]::ExtractToDirectory("$productZip", "$tempDir\", [System.Text.Encoding]::ascii) | Out-Null }
Catch { Write-Host "$error"}

Write-Host "Extracting $templateZip"
try { [System.IO.Compression.ZipFile]::ExtractToDirectory("$templateZip", "$tempDir\", [System.Text.Encoding]::ascii) | Out-Null }
Catch { Write-Host "$error"}

Write-host "==============="
Get-ChildItem -path $tempDir
Write-host "==============="

# Gather some path data
$productFolderPrefix = "$gitHubProductRepo"
$productExtracted = $(Get-ChildItem $tempDir -Directory -Recurse -Depth 0 | ? { $_.Name -match "^$productFolderPrefix" } | Sort LastWriteTime)[0].Name
$productLocation = "$tempDir\$productExtracted"
Write-Host "productLocation: $productLocation"

$templateFolderPrefix = "$githubTemplateOrg-$gitHubTemplateRepo"
$templateExtracted = $(Get-ChildItem $tempDir -Directory -Recurse -Depth 0 | ? { $_.Name -match "^$templateFolderPrefix" } | Sort LastWriteTime)[0].Name
$templateLocation = "$tempDir\$templateExtracted"
Write-Host "templateLocation: $templateLocation"


# Clean up the downloaded content
#Remove-Item -Path "$tempDir\$productZip" -Force
#Remove-Item -Path "$tempDir\$templateZip" -Force

#Run product, with transcript logging, and args passed from variables above
$script = "$productLocation\Set-OutlookSignatures.ps1"
If (Test-Path `$templateLocation\Signatures\variables.ps1` ) {
    powershell.exe -command "$script -graphonly $graphOnly -SignatureTemplatePath '$templateLocation\Signatures' -SignatureIniPath '$templateLocation\Signatures\_Signatures.ini' -ReplacementVariableConfigFile '$templateLocation\Signatures\variables.ps1' -SetCurrentUserOOFMessage $SetOofMsg -CreateRtfSignatures $CreateRtfSignatures -CreateTxtSignatures $CreateTxtSignatures -SignaturesForAutomappedAndAdditionalMailboxes $SignaturesForAutomappedAndAdditionalMailboxes -SetCurrentUserOutlookWebSignature $SetCurrentUserOutlookWebSignature -DeleteUserCreatedSignatures $DeleteUserCreatedSignatures -DeleteScriptCreatedSignaturesWithoutTemplate $DeleteScriptCreatedSignaturesWithoutTemplate"
}
Else {
    powershell.exe -command "$script -graphonly $graphOnly -SignatureTemplatePath '$templateLocation\Signatures' -SignatureIniPath '$templateLocation\Signatures\_Signatures.ini' -SetCurrentUserOOFMessage $SetOofMsg -CreateRtfSignatures $CreateRtfSignatures -CreateTxtSignatures $CreateTxtSignatures -SignaturesForAutomappedAndAdditionalMailboxes $SignaturesForAutomappedAndAdditionalMailboxes -SetCurrentUserOutlookWebSignature $SetCurrentUserOutlookWebSignature -DeleteUserCreatedSignatures $DeleteUserCreatedSignatures -DeleteScriptCreatedSignaturesWithoutTemplate $DeleteScriptCreatedSignaturesWithoutTemplate"

}
Stop-Transcript