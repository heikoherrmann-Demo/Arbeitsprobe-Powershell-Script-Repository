#Requires -Version 5.1
# ########################################
# Upload a File to the Repository
# ########################################
function Read-HostWithDefault {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [string]$Default
    )
    if ($Default) {
        $value = Read-Host "$Prompt [$Default]"
        if ([string]::IsNullOrWhiteSpace($value)) { return $Default }
        return $value.Trim()
    }
    do {
        $value = Read-Host $Prompt
    } while ([string]::IsNullOrWhiteSpace($value))
    return $value.Trim()
}


$UploadScript = Read-HostWithDefault -Prompt "Script Path & Name (C:\tmp\example.ps1)" -Default ""

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$Description = "Upload a PS1 File to the Repository"

Write-Host "Upload a Script" -ForegroundColor Green
Write-Host "  Script: $UploadScript" -ForegroundColor Green
Write-Host "  Descr.: $Description" -ForegroundColor Green

& "$scriptPath\PS_Upload.ps1" `
    -FilePath $UploadScript `
    -Description $Description
