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

$Scriptname = Read-HostWithDefault -Prompt "Scriptname without .ps1 (CTRL-C to abort)" -Default ""

Write-host "Execute: $Scriptname" -ForegroundColor Yellow
& "$scriptPath\PS_Execute.ps1" `
    -ScriptName $Scriptname `
    -ShowOutput