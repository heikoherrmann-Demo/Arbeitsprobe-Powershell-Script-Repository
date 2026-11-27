#Requires -Version 5.1
# ########################################
# Download a File from the Repository
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
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

# ########################################
# List exisiting Scripts
Write-host "Liste Scripts in Database:" -ForegroundColor Yellow
& "$scriptPath\PS_Execute.ps1" `
    -ListScripts

# #######################################
# Inputs
$TargetPath = Read-HostWithDefault -Prompt "Target Folder" -Default "C:\tmp\"
$Scriptname = Read-HostWithDefault -Prompt "Scriptname to Download without .ps1 (CTRL-C to abort)" -Default ""

& "$scriptPath\PS_Download.ps1" `
    -ScriptName $Scriptname `
    -TargetPath $TargetPath

# Advanced Techniques:
# --------------------
# With this the PS_Download.ps1 is in the Database and used directly from it:
# (needed to be uploaded first, but needs more time)
#
#& "$scriptPath\PS_Execute.ps1" `
#    -SqlServer $sqlServer -ScriptName PS_Download -ShowOutput `
#    -ScriptParameters @{ ScriptName = $Script; TargetPath = $TargetPath }