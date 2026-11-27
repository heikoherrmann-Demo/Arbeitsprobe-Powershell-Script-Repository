#Requires -Version 5.1
# ########################################
# Execute a PS1. Script from the Repository
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

# ########################################
# Execute a exisiting Scripts

$Scriptname1 = Read-HostWithDefault -Prompt "Scriptname1 without .ps1 (CTRL-C to abort)" -Default ""
$Scriptname2 = Read-HostWithDefault -Prompt "Scriptname2 without .ps1 (CTRL-C to abort)" -Default ""

Write-host "Execute: $Scriptname" -ForegroundColor Yellow
.\PS_Execute.ps1 -ScriptName $Scriptname1,$Scriptname2 -ScriptParameters @{ Mode = 'Full' } -PrintOutput

# Only the second Script ($Scriptname2) geht the Parameter 'Mode'. When non is needed :
# .\PS_Execute.ps1 -ScriptName $Scriptname1,$Scriptname2 -PrintOutput