<#
.SYNOPSIS
    Lädt Scripte aus dem zentralen Repository (SQL Server) in ein lokales Verzeichnis herunter.

.DESCRIPTION
    Das Script wird selbst im Repository gespeichert (Upload mit PS_Upload.ps1) und über
    PS_Execute.ps1 ausgeführt. Es läuft im Runspace von PS_Execute mit demselben Windows-Konto.

    Ablauf je Script:
      1. Script über dbo.PS_GetScript abrufen (der Abruf wird in PS_ScriptAccessLog protokolliert).
      2. SHA-256-Hash prüfen (UTF-16LE, identisch zu SQL Server). Bei Abweichung wird nichts geschrieben.
      3. Datei "<Scriptname>.ps1" im Zielverzeichnis anlegen (UTF-8 mit BOM, ohne Änderung der Zeilenenden).
      4. Geschriebene Datei zurücklesen und den Hash erneut prüfen.

    Schutz lokaler Änderungen:
      - Ist die Datei bereits vorhanden und inhaltlich identisch, wird nichts geschrieben (Unchanged).
      - Unterscheidet sich die vorhandene Datei, wird sie NICHT überschrieben (SkippedExists, Warnung),
        außer mit -Force.

    Statusergebnis in PS_Execute:
      SUCCESS  alle gewünschten Scripte sind heruntergeladen oder bereits aktuell
      WARNING  mindestens eine vorhandene abweichende Datei wurde nicht überschrieben
      ERROR    Script nicht gefunden, Hash-Abweichung, Schreibfehler oder grundlegender Fehler

    Aufruf über PS_Execute (die Parameter werden mit -ScriptParameters übergeben):
      .\PS_Execute.ps1 -SqlServer SQL01 -ScriptName PS_Download -ShowOutput `
          -ScriptParameters @{ SqlServer = 'SQL01'; ScriptName = 'Get-DiskReport'; TargetPath = 'C:\Scripts' }

      # Mehrere Scripte bzw. Platzhalter
      -ScriptParameters @{ SqlServer = 'SQL01'; ScriptName = 'Get-DiskReport','Hausmeister_*'; TargetPath = 'C:\Scripts' }

      # Alle aktiven Scripte, vorhandene abweichende Dateien überschreiben
      -ScriptParameters @{ SqlServer = 'SQL01'; All = $true; Force = $true; TargetPath = 'C:\Scripts' }

.PARAMETER SqlServer
    SQL-Server-Instanz (dieselbe wie bei PS_Execute).

.PARAMETER Database
    Datenbankname. Standard: PowerShellAutomation.

.PARAMETER ScriptName
    Name(n) der Scripte im Repository. Platzhalter (*, ?) sind erlaubt.

.PARAMETER All
    Alle aktiven Scripte herunterladen.

.PARAMETER TargetPath
    Zielverzeichnis (absoluter Pfad auf dem Rechner, auf dem PS_Execute läuft). Wird bei Bedarf
    angelegt. Das ausführende Konto benötigt Schreibrechte.

.PARAMETER Force
    Vorhandene, abweichende Dateien überschreiben.

.PARAMETER NoBom
    Dateien als UTF-8 ohne BOM schreiben (Standard: mit BOM, damit Windows PowerShell 5.1
    Umlaute korrekt liest).

.PARAMETER ValidateServerCertificate
    Zertifikatsprüfung aktivieren (Standard: aus, interner geschützter Betrieb).

.NOTES
    Konto benötigt: PS_ExecutorRole (PS_GetScripts, PS_GetScript).
    Es wird immer die aktuelle Version eines aktiven Scripts heruntergeladen.
    Achtung: Der Sourcecode liegt danach als Datei im Zielverzeichnis. Verzeichnis entsprechend schützen.
#>
[CmdletBinding(DefaultParameterSetName = 'ByName')]
param(
    [string]$SqlServer = '#DBSERVER#',
    [string]$Database = '#DBNAME#',

    [Parameter(Mandatory, ParameterSetName = 'ByName')]
    [ValidateNotNullOrEmpty()]
    [string[]]$ScriptName,

    [Parameter(Mandatory, ParameterSetName = 'All')]
    [switch]$All,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$TargetPath,

    [switch]$Force,
    [switch]$NoBom,
    [switch]$ValidateServerCertificate,

    [ValidateRange(5, 600)]
    [int]$CommandTimeoutSec = 60
)

$ErrorActionPreference = 'Stop'

# ======================================================================
# Hilfsfunktionen
# ======================================================================
function Get-SourceHash {
    # SQL Server berechnet den Hash über die NVARCHAR-Darstellung (UTF-16LE, ohne BOM).
    param([Parameter(Mandatory)][string]$Source)
    $bytes = [System.Text.Encoding]::Unicode.GetBytes($Source)
    $sha   = [System.Security.Cryptography.SHA256]::Create()
    try { [System.BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '') }
    finally { $sha.Dispose() }
}

function Get-SafeFileName {
    param([Parameter(Mandatory)][string]$Name)
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $chars = foreach ($c in $Name.ToCharArray()) { if ($invalid -contains $c) { '_' } else { [string]$c } }
    $safe = (-join $chars).TrimEnd('.', ' ')
    if ([string]::IsNullOrWhiteSpace($safe)) { throw "Aus dem Scriptnamen '$Name' lässt sich kein gültiger Dateiname bilden." }
    $safe
}

function New-ProcCommand {
    param([Parameter(Mandatory)][string]$Name)
    $cmd = New-Object "$($script:Provider).SqlCommand" $Name, $script:Conn
    $cmd.CommandType    = [System.Data.CommandType]::StoredProcedure
    $cmd.CommandTimeout = $CommandTimeoutSec
    $cmd
}

function Add-P {
    # $Type wird als Text übergeben (z. B. 'NVarChar') und zur Laufzeit in SqlDbType umgewandelt.
    param(
        [Parameter(Mandatory)]$Command,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Type,
        $Value,
        [int]$Size = 0
    )
    $dbType = [System.Data.SqlDbType]$Type
    if ($Size -ne 0) { $p = $Command.Parameters.Add($Name, $dbType, $Size) }
    else { $p = $Command.Parameters.Add($Name, $dbType) }
    $p.Value = if ($null -eq $Value) { [System.DBNull]::Value } else { $Value }
}

function Get-RepositoryScripts {
    $cmd  = New-ProcCommand -Name 'dbo.PS_GetScripts'
    $list = [System.Collections.Generic.List[object]]::new()
    try {
        Add-P -Command $cmd -Name '@ComputerName' -Type NVarChar -Size 256 -Value ([Environment]::MachineName)
        $reader = $cmd.ExecuteReader()
        try {
            while ($reader.Read()) {
                $list.Add([pscustomobject]@{
                        ScriptID      = [int]$reader['ScriptID']
                        ScriptName    = [string]$reader['ScriptName']
                        ScriptVersion = [int]$reader['ScriptVersion']
                    })
            }
        }
        finally { $reader.Dispose() }
    }
    finally { $cmd.Dispose() }
    $list.ToArray()    # einzelne Objekte ausgeben (Pipeline entpackt das Array)
}

function Get-RepositoryScript {
    param([Parameter(Mandatory)][int]$Id)
    $cmd = New-ProcCommand -Name 'dbo.PS_GetScript'
    try {
        Add-P -Command $cmd -Name '@ScriptID' -Type Int -Value $Id
        $reader = $cmd.ExecuteReader()
        try {
            if (-not $reader.Read()) { throw "PS_GetScript hat für ScriptID $Id kein Ergebnis geliefert." }
            [pscustomobject]@{
                ScriptID      = [int]$reader['ScriptID']
                ScriptName    = [string]$reader['ScriptName']
                ScriptVersion = [int]$reader['ScriptVersion']
                ScriptSource  = [string]$reader['ScriptSource']
                HashSHA256    = [string]$reader['HashSHA256']
            }
        }
        finally { $reader.Dispose() }
    }
    finally { $cmd.Dispose() }
}

# ======================================================================
# Hauptablauf
# ======================================================================
$script:Conn     = $null
$script:Provider = $null

try {
    if (-not [System.IO.Path]::IsPathRooted($TargetPath)) {
        throw "TargetPath muss ein absoluter Pfad sein: '$TargetPath'"
    }

    # --- Zielverzeichnis ---
    if (Test-Path -LiteralPath $TargetPath) {
        if (-not (Test-Path -LiteralPath $TargetPath -PathType Container)) {
            throw "TargetPath '$TargetPath' ist kein Verzeichnis."
        }
    }
    else {
        $null = New-Item -ItemType Directory -Path $TargetPath -Force
        Write-Host "Zielverzeichnis angelegt: $TargetPath"
    }

    # --- SQL-Verbindung ---
    if ($PSVersionTable.PSVersion.Major -lt 6) { Add-Type -AssemblyName System.Data }
    foreach ($ns in 'System.Data.SqlClient', 'Microsoft.Data.SqlClient') {
        if ("$ns.SqlConnection" -as [type]) { $script:Provider = $ns; break }
    }
    if (-not $script:Provider) { throw 'Kein SQL-Client gefunden (System.Data.SqlClient bzw. Microsoft.Data.SqlClient).' }

    $connString = 'Server={0};Database={1};Integrated Security=True;Encrypt=True;TrustServerCertificate={2};Application Name=PS_Download;Connect Timeout=15' -f `
        $SqlServer, $Database, (-not $ValidateServerCertificate.IsPresent)
    $script:Conn = New-Object "$($script:Provider).SqlConnection" $connString
    $script:Conn.Open()

    # --- Auswahl der Scripte ---
    $available = @(Get-RepositoryScripts)
    if ($All) {
        $selected = $available
    }
    else {
        $selected = @()
        foreach ($pattern in $ScriptName) {
            $found = @($available | Where-Object { $_.ScriptName -like $pattern })
            if ($found.Count -eq 0) {
                Write-Error -Message "Script '$pattern' wurde nicht gefunden oder ist nicht aktiv." -ErrorAction Continue
                [pscustomobject]@{ ScriptName = $pattern; Version = $null; Path = $null; Result = 'NotFound'; Hash = $null }
            }
            else { $selected += $found }
        }
        $selected = @($selected | Sort-Object -Property ScriptID -Unique)
    }

    Write-Host "$($selected.Count) Script(e) ausgewählt, Ziel: $TargetPath"

    $enc = New-Object System.Text.UTF8Encoding(-not $NoBom.IsPresent)
    $readEnc = New-Object System.Text.UTF8Encoding($false)    # beim Lesen erkennt .NET vorhandene BOMs selbst
    $counts = @{ Downloaded = 0; Overwritten = 0; Unchanged = 0; SkippedExists = 0; Failed = 0 }

    foreach ($item in $selected) {
        $result = [pscustomobject]@{
            ScriptName = $item.ScriptName
            Version    = $item.ScriptVersion
            Path       = $null
            Result     = 'Failed'
            Hash       = $null
        }
        $tmp = $null

        try {
            $rec = Get-RepositoryScript -Id $item.ScriptID
            $result.Version = $rec.ScriptVersion

            $localHash = Get-SourceHash -Source $rec.ScriptSource
            if ($localHash -ne $rec.HashSHA256) {
                throw "Hash-Abweichung (berechnet=$localHash, Datenbank=$($rec.HashSHA256)). Es wurde nichts geschrieben."
            }
            $result.Hash = $localHash

            $fileName = (Get-SafeFileName -Name $rec.ScriptName) + '.ps1'
            $path = Join-Path $TargetPath $fileName
            $result.Path = $path

            $action = 'Write'
            $exists = Test-Path -LiteralPath $path -PathType Leaf
            if ($exists) {
                $existingText = [System.IO.File]::ReadAllText($path, $readEnc)
                if ((Get-SourceHash -Source $existingText) -eq $localHash) { $action = 'Unchanged' }
                elseif ($Force) { $action = 'Overwrite' }
                else { $action = 'Skip' }
            }

            switch ($action) {
                'Unchanged' {
                    $result.Result = 'Unchanged'
                    $counts.Unchanged++
                    Write-Host "'$($rec.ScriptName)': bereits aktuell (Version $($rec.ScriptVersion))."
                }
                'Skip' {
                    $result.Result = 'SkippedExists'
                    $counts.SkippedExists++
                    Write-Warning "'$path' existiert und weicht von Version $($rec.ScriptVersion) ab. Nicht überschrieben (zum Überschreiben -Force angeben)."
                }
                default {
                    # Atomar schreiben: erst temporäre Datei, dann umbenennen
                    $tmp = Join-Path $TargetPath ('.{0}.{1}.tmp' -f $fileName, [guid]::NewGuid().ToString('N'))
                    [System.IO.File]::WriteAllText($tmp, $rec.ScriptSource, $enc)
                    Move-Item -LiteralPath $tmp -Destination $path -Force
                    $tmp = $null

                    # Kontrolle: geschriebene Datei lesen und Hash prüfen
                    $written = [System.IO.File]::ReadAllText($path, $readEnc)
                    if ((Get-SourceHash -Source $written) -ne $localHash) {
                        throw "Die geschriebene Datei '$path' weicht vom Sourcecode in der Datenbank ab."
                    }

                    if ($action -eq 'Overwrite') { $result.Result = 'Overwritten'; $counts.Overwritten++ }
                    else { $result.Result = 'Downloaded'; $counts.Downloaded++ }
                    Write-Host "'$($rec.ScriptName)': Version $($rec.ScriptVersion) -> $path ($($result.Result))."
                }
            }
        }
        catch {
            $result.Result = 'Failed'
            $counts.Failed++
            Write-Error -Message "Download von '$($item.ScriptName)' fehlgeschlagen: $($_.Exception.GetBaseException().Message)" -ErrorAction Continue
        }
        finally {
            if ($tmp -and (Test-Path -LiteralPath $tmp)) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
        }

        $result
    }

    Write-Host ("Fertig: {0} heruntergeladen, {1} überschrieben, {2} bereits aktuell, {3} übersprungen, {4} fehlgeschlagen." -f `
            $counts.Downloaded, $counts.Overwritten, $counts.Unchanged, $counts.SkippedExists, $counts.Failed)
}
finally {
    if ($null -ne $script:Conn) { try { $script:Conn.Dispose() } catch { } }
}
