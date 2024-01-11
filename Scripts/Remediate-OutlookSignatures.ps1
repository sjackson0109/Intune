<#
.SYNOPSIS
    PowerShell script to remediate an pre/non existing Email Signatures; from Set-OutlookSignatures script.

.EXAMPLE
    .\Remediate-EmailSignatures.ps1

.DESCRIPTION
    This PowerShell script is deployed as a remediation script using Microsoft Intune remediations.

.LINK
    https://github.com/sjackson0109/Intune/blob/main/Device%20Remediation/Remediate-EmailSignatures.ps1

.LINK
    https://github.com/Set-OutlookSignatures/Set-OutlookSignatures

.LINK
    https://learn.microsoft.com/en-us/mem/intune/fundamentals/remediations

.NOTES
    Version:        1.0.3
    Creation Date:  2023-11-07
    Last Updated:   2023-12-22
    Author:         Simon Jackson / sjackson0109
#>
$temp = $(Get-Location).path
$logFile = "$temp\Set-OutlookSignatures.log"
Start-Transcript $logFile -Append

# Variables for Download and Extract
$githubProductOrg = "Set-OutlookSignatures"     # Author of the main script
$githubProductRepo = "Set-OutlookSignatures"    # Repo of the main script
$githubTemplateOrg = "sjackson0109"             # update this to your own org 
$githubTemplateRepo = "EmailTemplates"          # update this to your own repo, upload your templates into the `Templates` sub-folder, and create a release

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
$productRelease | Out-File "$githubProductRepo.json"
$productZip = "$githubProductRepo-$productVersion.zip"
$productPath = "$githubProductRepo-$productVersion" -replace '-v' , '_v'

$templateRelease | Out-File "$githubTemplateRepo.json"
$templateZip = "$githubTemplateRepo-$templateVersion.zip"
$templatePath = "$githubTemplateOrg-$githubTemplateRepo-$templateVersion" 

Add-Type -AssemblyName System.IO.Compression.FileSystem


# Check if the latest version is already downloaded, clean up the file-system and download+extract, or just extract again
If (Test-Path "$temp\$productPath"){
    Write-Host "Cleaning up local path $productPath"
    Remove-Item $productPath -recurse -Force
} else {
    Write-Host "Creating product folder"
    New-Item $productPath -ErrorAction SilentlyContinue
    Write-Host "Downloading $productUrl to $productZip"
    Invoke-WebRequest $productUrl -Out $productZip
}

If (Test-Path "$temp\$templatePath"){
    Write-Host "Cleaning up local path $templatePath"
    Remove-Item $templatePath -recurse -Force
} else {
    Write-Host "Downloading $templateUrl to $templateZip"
    Invoke-WebRequest $templateUrl -Out $templateZip
}
Write-host "===============" 
Get-ChildItem .\
Write-host "==============="


# A fresh Extraction of the zipball files to the temp directory, filename encoding needs converting to ascii, not utf8.
# Note: some errors with file-name length when testing with my user docs area. C:\WINDOWS\IMECache\HealthScripts\(GUID)\ is just as long, so skip errors. Only signature samples anyway, don't need them.
Write-Host "Extracting $productZip"
[System.IO.Compression.ZipFile]::ExtractToDirectory("$temp\$productZip", "$temp\", [System.Text.Encoding]::ascii)
Write-Host "Extracting $templateZip"
[System.IO.Compression.ZipFile]::ExtractToDirectory("$temp\$templateZip", "$temp\", [System.Text.Encoding]::ascii)
#Expand-Archive -path "$temp\$templateZip" -destinationPath "$temp\$targetPath" | out-null

Write-host "==============="
Get-ChildItem $temp
Write-host "==============="

# Gather some path data
$templateFolderPrefix = "$githubTemplateOrg-$gitHubTemplateRepo"
$templateLocation = $(Get-ChildItem $temp -Directory -Recurse -Depth 0 | ? { $_.Name -match "^$templateFolderPrefix" } | Sort LastWriteTime)[0].Name
$templateLocation = "$temp\$templateExtracted"

# Clean up the downloaded content
#Remove-Item -Path $productZip -Force
#Remove-Item -Path $templateZip -Force


#Run product, with transcript logging, and args passed from variables above
#Set-Location "$temp\$productPath"
$script = ".\$productPath\Set-OutlookSignatures.ps1"
powershell.exe -command "$script -graphonly $graphOnly -SignatureTemplatePath '$($(Get-Location).path)\$templateLocation\Signatures' -SignatureIniPath '$($(Get-Location).path)\$templateLocation\Signatures\_Signatures.ini' -SetCurrentUserOOFMessage $SetOofMsg -CreateRtfSignatures $CreateRtfSignatures -CreateTxtSignatures $CreateTxtSignatures -SignaturesForAutomappedAndAdditionalMailboxes $SignaturesForAutomappedAndAdditionalMailboxes -SetCurrentUserOutlookWebSignature $SetCurrentUserOutlookWebSignature -DeleteUserCreatedSignatures $DeleteUserCreatedSignatures -DeleteScriptCreatedSignaturesWithoutTemplate $DeleteScriptCreatedSignaturesWithoutTemplate"
Stop-Transcript
exit 0