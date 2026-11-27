#Requires -Version 5.1
<#
.SYNOPSIS
    Ruft ein PowerShell-Script aus dem zentralen Repository (SQL Server) ab, prüft den
    SHA-256-Hash, führt es lokal aus und überträgt Protokoll und Ausgaben zurück in die Datenbank.

.DESCRIPTION
    Ablauf eines Laufs:
      1. Verbindung zum SQL Server (Windows Integrated Authentication, Zertifikatsprüfung aus).
      2. Zwischengespeicherte Ergebnisse früherer Läufe nachsenden (siehe Offline-Verhalten).
      3. Script ermitteln (Name oder ID) und über dbo.PS_GetScript abrufen (der Abruf wird
         serverseitig in PS_ScriptAccessLog protokolliert).
      4. SHA-256-Hash lokal berechnen (UTF-16LE, identisch zu SQL Server) und mit dem Hash der
         Datenbank vergleichen. Bei Abweichung wird NICHT ausgeführt.
      5. Systeminformationen erfassen (Rechner, ausführendes Konto, PowerShell-Version/-Edition,
         Betriebssystem, installierte Module mit Version und Pfad).
      6. Ausführungsbeginn protokollieren (Status STARTED), Script ausführen, Ausgaben und
         Fehler einsammeln, Ergebnis protokollieren (SUCCESS, WARNING, ERROR oder CANCELLED).
      7. Ausgaben werden über dieselbe ExecutionID in PS_ExecutionOutput gespeichert.

    Ausführung: Das Script wird NICHT als Datei auf die Festplatte geschrieben, sondern in einem
    eigenen Runspace innerhalb dieses Prozesses ausgeführt (gleiches Konto, saubere Variablen,
    unabhängig von der Execution Policy). Hinweise für Scripte im Repository:
      - Fehler mit throw melden; "exit" vermeiden, Read-Host und andere Eingaben sind nicht möglich.
      - $PSScriptRoot und $MyInvocation.MyCommand.Path sind nicht belegt (es gibt keine Datei).
      - Parameter des Scripts werden über -ScriptParameters übergeben.
      - Im Runspace steht die Variable $PS_ExecutionID (GUID dieses Laufs) zur Verfügung,
        z. B. um eigene Ergebnisse mit PS_ExecutionLog zu verknüpfen.

    Mehrere Scripte in einer Session:
      Mit -ScriptName A,B,C werden mehrere Scripte nacheinander im SELBEN Runspace ausgeführt. Funktionen,
      Variablen und geladene Module früherer Scripte stehen den späteren zur Verfügung. Beispiel: Lib_Common
      definiert Funktionen, Main_Job ruft sie auf (-ScriptName Lib_Common,Main_Job). Regeln:
        - Das LETZTE Script ist das Hauptscript: Es erhält -ScriptParameters, und der Lauf wird unter seinem
          Namen protokolliert (PS_ExecutionLog, -ShowLastRun, -ShowOutput). Alle beteiligten Scripte mit
          Version stehen als erster Ausgabeeintrag (INFORMATION); die Ausgaben tragen den Scriptnamen als Präfix.
        - Alle Scripte werden zuerst abgerufen und per Hash geprüft. Ist eines nicht auffindbar oder
          manipuliert, wird KEIN Script ausgeführt. Jeder Abruf steht im Zugriffsprotokoll.
        - Bricht ein Script mit einem terminierenden Fehler (throw) ab oder läuft in den Timeout, werden die
          folgenden Scripte nicht ausgeführt. Nicht terminierende Fehler stoppen die Session nicht.
        - -TimeoutSec gilt für alle Scripte zusammen. Der Status ist der schlechteste aller Scripte.
      Namen werden durch Komma getrennt (bei Aufruf mit -File der Aufgabenplanung auch ohne Anführungszeichen).

    Statusermittlung:
      ERROR     terminierender Fehler oder mindestens ein Eintrag im Fehlerstrom
      WARNING   keine Fehler, aber mindestens eine Warnung
      SUCCESS   weder Fehler noch Warnungen
      CANCELLED Zeitüberschreitung (-TimeoutSec) oder Abbruch (z. B. Strg+C)

    Offline-Verhalten: Ist die Datenbank nach der Ausführung nicht erreichbar, werden Protokoll und
    Ausgaben lokal in einer Warteschlange gespeichert (DPAPI-verschlüsselt, an das ausführende Konto
    gebunden, Verzeichnis nur für dieses Konto, SYSTEM und Administratoren zugänglich) und beim
    nächsten Lauf bzw. mit -FlushQueue übertragen. Läuft die Verbindung bereits zu Beginn nicht,
    kann das Script nicht abgerufen werden; der Lauf endet mit Exitcode 2.

    Abfragemodus (führt nichts aus, zeigt nur gespeicherte Ergebnisse):
      -ShowOutput  -ScriptName <Name>   zeigt die gespeicherten Ausgaben des letzten Laufs
      -ShowLastRun -ScriptName <Name>   zeigt Status, Zeiten, Dauer und Fehlermeldung des letzten Laufs
    Mit -ComputerName <Rechner> wird auf einen Rechner eingeschränkt, sonst gilt der zuletzt
    gestartete Lauf auf beliebigem Rechner. Benötigt PS_ViewerRole (siehe SQL_ShowLastRun.sql).
    Zeitangaben werden in der lokalen Zeit dieses Rechners angezeigt.

.PARAMETER ScriptName
    Name des Scripts im Repository (entspricht dem Dateinamen ohne .ps1 beim Upload). Mehrere Namen,
    durch Komma getrennt, werden in dieser Reihenfolge in einer Session ausgeführt (das letzte ist das
    Hauptscript). Scriptnamen dürfen dafür kein Komma enthalten.

.PARAMETER ScriptID
    Alternativ zum Namen: ID eines einzelnen Scripts (für mehrere Scripte -ScriptName verwenden).

.PARAMETER ListScripts
    Listet die für diesen Rechner verfügbaren (aktiven) Scripte auf und beendet sich.

.PARAMETER FlushQueue
    Überträgt nur die lokal zwischengespeicherten Ergebnisse und beendet sich.

.PARAMETER SqlServer
    SQL-Server-Instanz, z. B. "SQL01" oder "SQL01\INST1".

.PARAMETER Database
    Datenbankname. Standard: PowerShellAutomation.

.PARAMETER ScriptParameters
    Hashtable mit Parametern, die an das Repository-Script übergeben werden. Bei mehreren Scripten
    erhält nur das letzte (Hauptscript) die Parameter.

.PARAMETER TimeoutSec
    Maximale Laufzeit in Sekunden (0 = unbegrenzt). Standard: 3600. Bei mehreren Scripten gilt die
    Zeit für alle zusammen. Bei Überschreitung wird das laufende Script gestoppt, der Status ist CANCELLED.

.PARAMETER CaptureVerbose
    Verbose-Ausgaben des Scripts erfassen (setzt $VerbosePreference auf Continue).

.PARAMETER CaptureDebug
    Debug-Ausgaben des Scripts erfassen (setzt $DebugPreference auf Continue).

.PARAMETER SkipModuleInventory
    Modulinventur überspringen (Get-Module -ListAvailable kann einige Sekunden dauern).

.PARAMETER ShowOutput
    Abfragemodus: zeigt die gespeicherten Ausgaben des letzten Laufs von -ScriptName (kein Ausführen).
    Bisher bedeutete -ShowOutput "Ausgaben des gerade ausgeführten Laufs anzeigen": Das heißt jetzt -PrintOutput.

.PARAMETER ShowLastRun
    Abfragemodus: zeigt Status, Start, Ende, Dauer, Rechner, Konto und Fehlermeldung des letzten Laufs von -ScriptName.

.PARAMETER ComputerName
    Nur für -ShowOutput und -ShowLastRun: letzten Lauf auf diesem Rechner suchen (Standard: beliebiger Rechner).

.PARAMETER OutputType
    Nur für -ShowOutput: nur diese Ausgabetypen anzeigen (OUTPUT, ERROR, WARNING, INFORMATION, VERBOSE, DEBUG).

.PARAMETER PrintOutput
    Erfasste Ausgaben des gerade ausgeführten Scripts zusätzlich auf der Konsole anzeigen.

.PARAMETER NoFlush
    Wartende Ergebnisse früherer Läufe zu Beginn nicht nachsenden.

.PARAMETER ValidateServerCertificate
    Standardmäßig wird das SQL-Server-Zertifikat NICHT geprüft (interner, geschützter Betrieb).
    Mit diesem Schalter wird die Prüfung aktiviert. Die Verbindung bleibt immer verschlüsselt.

.PARAMETER LogLevel
    Detailgrad der lokalen Logdatei: ERROR, WARNING, INFO (Standard) oder DEBUG.

.PARAMETER LogPath
    Verzeichnis der lokalen Logdatei. Standard: %LOCALAPPDATA%\PSAutomation\Logs

.PARAMETER QueuePath
    Verzeichnis der lokalen Warteschlange. Standard: %LOCALAPPDATA%\PSAutomation\Queue

.PARAMETER MaxOutputChars
    Maximale Zeichenzahl pro gespeichertem Ausgabeeintrag (Standard 1.000.000), längere Texte werden gekürzt.

.PARAMETER MaxOutputEntries
    Maximale Anzahl gespeicherter Ausgabeeinträge pro Lauf (Standard 2000), weitere werden verworfen.

.EXAMPLE
    .\PS_Execute.ps1 -ListScripts -SqlServer SQL01

.EXAMPLE
    .\PS_Execute.ps1 -ScriptName Get-DiskReport -SqlServer SQL01 -PrintOutput

.EXAMPLE
    # Abfragemodus: Ausgaben bzw. Status des letzten Laufs anzeigen (es wird nichts ausgeführt)
    .\PS_Execute.ps1 -ShowOutput  -ScriptName Get-DiskReport -SqlServer SQL01
    .\PS_Execute.ps1 -ShowLastRun -ScriptName Get-DiskReport -SqlServer SQL01 -ComputerName VM-SCRIPT
    .\PS_Execute.ps1 -ShowOutput  -ScriptName Get-DiskReport -SqlServer SQL01 -OutputType ERROR,WARNING

.EXAMPLE
    .\PS_Execute.ps1 -ScriptName Cleanup-Temp -SqlServer SQL01 -ScriptParameters @{ OlderThanDays = 30 }

.EXAMPLE
    # Mehrere Scripte in einer Session: Lib_Common (Funktionen), danach das Hauptscript Main_Job
    .\PS_Execute.ps1 -ScriptName Lib_Common,Lib_Logging,Main_Job -SqlServer SQL01 -ScriptParameters @{ Mode = 'Full' }

.EXAMPLE
    # Windows-Aufgabenplanung (Programm: pwsh.exe bzw. powershell.exe)
    -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\PS\PS_Execute.ps1" -SqlServer SQL01 -ScriptName Hausmeister_Cleanup

.NOTES
    Exitcodes: 0 = Erfolg (SUCCESS/WARNING), 1 = Script mit ERROR oder CANCELLED beendet,
               2 = keine SQL-Verbindung, 3 = Script nicht gefunden oder nicht abrufbar,
               4 = Hash-Abweichung (Script wurde nicht ausgeführt). Im Abfragemodus bedeutet 3:
               kein Lauf gefunden.
    Benötigt EXECUTE-Rechte auf PS_GetScripts, PS_GetScript, PS_WriteExecutionLog und
    PS_WriteExecutionOutput (PS_ExecutorRole). Der Abfragemodus benötigt zusätzlich PS_GetLastExecution
    und PS_GetExecutionOutput (PS_ViewerRole). Lauffähig unter Windows PowerShell 5.1 und PowerShell 7+.
#>
[CmdletBinding(DefaultParameterSetName = 'ByName')]
param(
    [Parameter(Mandatory, ParameterSetName = 'ByName', Position = 0)]
    [Parameter(Mandatory, ParameterSetName = 'ShowOutput', Position = 0)]
    [Parameter(Mandatory, ParameterSetName = 'ShowLastRun', Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string[]]$ScriptName,

    [Parameter(Mandatory, ParameterSetName = 'ById')]
    [ValidateRange(1, 2147483647)]
    [int]$ScriptID,

    [Parameter(Mandatory, ParameterSetName = 'List')]
    [switch]$ListScripts,

    [Parameter(Mandatory, ParameterSetName = 'Flush')]
    [switch]$FlushQueue,

    [Parameter(Mandatory, ParameterSetName = 'ShowOutput')]
    [switch]$ShowOutput,

    [Parameter(Mandatory, ParameterSetName = 'ShowLastRun')]
    [switch]$ShowLastRun,

    [Parameter(ParameterSetName = 'ShowOutput')]
    [Parameter(ParameterSetName = 'ShowLastRun')]
    [ValidateNotNullOrEmpty()]
    [string]$ComputerName,

    [Parameter(ParameterSetName = 'ShowOutput')]
    [ValidateSet('OUTPUT', 'ERROR', 'WARNING', 'INFORMATION', 'VERBOSE', 'DEBUG')]
    [string[]]$OutputType,

    
    [string]$SqlServer = '#DBSERVER#',
    [string]$Database = '#DBNAME#',

    [Parameter(ParameterSetName = 'ByName')]
    [Parameter(ParameterSetName = 'ById')]
    [hashtable]$ScriptParameters,

    [ValidateRange(0, 604800)]
    [int]$TimeoutSec = 3600,

    [switch]$CaptureVerbose,
    [switch]$CaptureDebug,
    [switch]$SkipModuleInventory,
    [switch]$PrintOutput,
    [switch]$NoFlush,
    [switch]$ValidateServerCertificate,

    [ValidateSet('ERROR', 'WARNING', 'INFO', 'DEBUG')]
    [string]$LogLevel = 'INFO',

    [string]$LogPath = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'PSAutomation\Logs'),
    [string]$QueuePath = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'PSAutomation\Queue'),

    [ValidateRange(1000, 100000000)]
    [int]$MaxOutputChars = 1000000,

    [ValidateRange(1, 100000)]
    [int]$MaxOutputEntries = 2000,

    [ValidateRange(5, 600)]
    [int]$CommandTimeoutSec = 60
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

# ---------- Zustand ----------
$script:SetName         = $PSCmdlet.ParameterSetName
$script:ExitCode        = 0
$script:Conn            = $null
$script:Provider        = $null
$script:ExecutionId     = $null
$script:Ctx             = $null
$script:Run             = $null
$script:LogFile         = $null
$script:DpapiAvailable  = $false
$script:DpapiEntropy    = [System.Text.Encoding]::UTF8.GetBytes('PS_Execute.Queue.v1')
$script:LevelRank       = @{ ERROR = 0; WARNING = 1; INFO = 2; DEBUG = 3 }

# ======================================================================
# Logging (lokale Logdatei)
# ======================================================================
function Write-ExecLog {
    param(
        [Parameter(Position = 0)]
        [ValidateSet('ERROR', 'WARNING', 'INFO', 'DEBUG')]
        [string]$Level = 'INFO',

        [Parameter(Mandatory, Position = 1)]
        [string]$Message
    )
    if ($script:LevelRank[$Level] -gt $script:LevelRank[$LogLevel]) { return }

    $idPart = if ($null -ne $script:ExecutionId) { " [$($script:ExecutionId)]" } else { '' }
    $line = '{0:yyyy-MM-dd HH:mm:ss.fff}Z [{1}] [{2}\{3}@{4}]{5} {6}' -f `
        [datetime]::UtcNow, $Level, [Environment]::UserDomainName, [Environment]::UserName,
        [Environment]::MachineName, $idPart, $Message

    if ($script:LogFile) {
        try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding utf8 }
        catch { Write-Warning "Logdatei nicht beschreibbar: $($_.Exception.Message)" }
    }

    switch ($Level) {
        'ERROR'   { Write-Host $Message -ForegroundColor Red }
        'WARNING' { Write-Warning $Message }
        'DEBUG'   { Write-Verbose $Message }
        default   { Write-Host $Message }
    }
}

# ======================================================================
# Hilfsfunktionen
# ======================================================================
function Get-ErrorText {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)
    $ErrorRecord.Exception.GetBaseException().Message
}

function Get-SourceHash {
    # SQL Server berechnet den Hash über die NVARCHAR-Darstellung (UTF-16LE, ohne BOM).
    param([Parameter(Mandatory)][string]$Source)
    $bytes = [System.Text.Encoding]::Unicode.GetBytes($Source)
    $sha   = [System.Security.Cryptography.SHA256]::Create()
    try { [System.BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '') }
    finally { $sha.Dispose() }
}

function Format-ErrorRecord {
    param($Record)
    $text = $Record.ToString()
    $inv  = $Record.InvocationInfo
    if ($null -ne $inv -and -not [string]::IsNullOrWhiteSpace($inv.PositionMessage)) {
        $text += [Environment]::NewLine + $inv.PositionMessage.Trim()
    }
    $text
}

function Get-SystemInfo {
    $runsAs = try { [System.Security.Principal.WindowsIdentity]::GetCurrent().Name }
              catch { '{0}\{1}' -f [Environment]::UserDomainName, [Environment]::UserName }

    $os = try {
        $c = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        '{0} {1} (Build {2})' -f $c.Caption.Trim(), $c.Version, $c.BuildNumber
    }
    catch { [Environment]::OSVersion.VersionString }
    if ($os.Length -gt 500) { $os = $os.Substring(0, 500) }

    $modulesJson = $null
    if (-not $SkipModuleInventory) {
        try {
            $mods = @(Get-Module -ListAvailable -ErrorAction SilentlyContinue |
                    Sort-Object -Property Name, Version |
                    ForEach-Object {
                        [pscustomobject]@{
                            Name    = $_.Name
                            Version = $_.Version.ToString()
                            Path    = $_.ModuleBase
                        }
                    })
            $modulesJson = ConvertTo-Json -InputObject $mods -Compress
            Write-ExecLog -Level DEBUG -Message "Modulinventur: $($mods.Count) Module erfasst."
        }
        catch {
            Write-ExecLog -Level WARNING -Message "Modulinventur fehlgeschlagen: $(Get-ErrorText $_)"
        }
    }

    [pscustomobject]@{
        ComputerName      = [Environment]::MachineName
        RunsAs            = $runsAs
        PowerShellVersion = $PSVersionTable.PSVersion.ToString()
        PowerShellEdition = $PSVersionTable.PSEdition
        OSVersion         = $os
        InstalledModules  = $modulesJson
    }
}

# ======================================================================
# SQL-Zugriff
# ======================================================================
function Open-SqlConnection {
    if ($null -ne $script:Conn) {
        try { $script:Conn.Dispose() } catch { }
        $script:Conn = $null
    }
    $connString = 'Server={0};Database={1};Integrated Security=True;Encrypt=True;TrustServerCertificate={2};Application Name=PS_Execute.ps1;Connect Timeout=15' -f `
        $SqlServer, $Database, (-not $ValidateServerCertificate.IsPresent)

    $conn = New-Object "$($script:Provider).SqlConnection" $connString
    try { $conn.Open() }
    catch { $conn.Dispose(); throw }
    $script:Conn = $conn
}

function New-SqlCommandObject {
    param([Parameter(Mandatory)][string]$ProcedureName)
    $cmd = New-Object "$($script:Provider).SqlCommand" $ProcedureName, $script:Conn
    $cmd.CommandType    = [System.Data.CommandType]::StoredProcedure
    $cmd.CommandTimeout = $CommandTimeoutSec
    $cmd
}

function Add-SqlParameter {
    # $Type wird bewusst als Text übergeben (z. B. 'NVarChar') und zur Laufzeit in SqlDbType umgewandelt.
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
    $cmd  = New-SqlCommandObject -ProcedureName 'dbo.PS_GetScripts'
    $list = [System.Collections.Generic.List[object]]::new()
    try {
        Add-SqlParameter -Command $cmd -Name '@ComputerName' -Type NVarChar -Size 256 -Value ([Environment]::MachineName)
        $reader = $cmd.ExecuteReader()
        try {
            while ($reader.Read()) {
                $list.Add([pscustomobject]@{
                        ScriptID      = [int]$reader['ScriptID']
                        ScriptName    = [string]$reader['ScriptName']
                        Description   = if ($reader['Description'] -is [System.DBNull]) { $null } else { [string]$reader['Description'] }
                        ScriptVersion = [int]$reader['ScriptVersion']
                        HashSHA256    = [string]$reader['HashSHA256']
                        ModifiedAt    = [datetime]$reader['ModifiedAt']
                    })
            }
        }
        finally { $reader.Dispose() }
    }
    finally { $cmd.Dispose() }
    # Einzelne Objekte ausgeben (kein vorangestelltes Komma), damit die Pipeline sie entpackt.
    $list.ToArray()
}

function Get-RepositoryScript {
    param([Parameter(Mandatory)][int]$Id)
    $cmd = New-SqlCommandObject -ProcedureName 'dbo.PS_GetScript'
    try {
        Add-SqlParameter -Command $cmd -Name '@ScriptID' -Type Int -Value $Id
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

function Write-ExecutionLogToDb {
    param([Parameter(Mandatory)]$Entry)
    $cmd = New-SqlCommandObject -ProcedureName 'dbo.PS_WriteExecutionLog'
    try {
        $finished = if ($null -ne $Entry.FinishedTicks) { [datetime]::new([long]$Entry.FinishedTicks, [System.DateTimeKind]::Utc) } else { $null }
        $duration = if ($null -ne $Entry.DurationMs) { [long]$Entry.DurationMs } else { $null }

        Add-SqlParameter -Command $cmd -Name '@ExecutionID'       -Type UniqueIdentifier -Value ([guid]$Entry.ExecutionID)
        Add-SqlParameter -Command $cmd -Name '@ScriptID'          -Type Int              -Value ([int]$Entry.ScriptID)
        Add-SqlParameter -Command $cmd -Name '@ScriptVersion'     -Type Int              -Value ([int]$Entry.ScriptVersion)
        Add-SqlParameter -Command $cmd -Name '@ComputerName'      -Type NVarChar -Size 256 -Value $Entry.ComputerName
        Add-SqlParameter -Command $cmd -Name '@RunsAs'            -Type NVarChar -Size 256 -Value $Entry.RunsAs
        Add-SqlParameter -Command $cmd -Name '@PowerShellVersion' -Type NVarChar -Size 50  -Value $Entry.PowerShellVersion
        Add-SqlParameter -Command $cmd -Name '@PowerShellEdition' -Type NVarChar -Size 50  -Value $Entry.PowerShellEdition
        Add-SqlParameter -Command $cmd -Name '@OSVersion'         -Type NVarChar -Size 500 -Value $Entry.OSVersion
        Add-SqlParameter -Command $cmd -Name '@InstalledModules'  -Type NVarChar -Size -1  -Value $Entry.InstalledModules
        Add-SqlParameter -Command $cmd -Name '@ExecutionStatus'   -Type VarChar  -Size 20  -Value $Entry.ExecutionStatus
        Add-SqlParameter -Command $cmd -Name '@StartedAt'         -Type DateTime2 -Value ([datetime]::new([long]$Entry.StartedTicks, [System.DateTimeKind]::Utc))
        Add-SqlParameter -Command $cmd -Name '@FinishedAt'        -Type DateTime2 -Value $finished
        Add-SqlParameter -Command $cmd -Name '@DurationMs'        -Type BigInt    -Value $duration
        Add-SqlParameter -Command $cmd -Name '@ErrorMessage'      -Type NVarChar -Size -1  -Value $Entry.ErrorMessage
        $null = $cmd.ExecuteNonQuery()
    }
    finally { $cmd.Dispose() }
}

function Write-ExecutionOutputToDb {
    param(
        [Parameter(Mandatory)][string]$ExecutionId,
        [Parameter(Mandatory)][string]$OutputType,
        [string]$OutputText
    )
    $cmd = New-SqlCommandObject -ProcedureName 'dbo.PS_WriteExecutionOutput'
    try {
        Add-SqlParameter -Command $cmd -Name '@ExecutionID' -Type UniqueIdentifier -Value ([guid]$ExecutionId)
        Add-SqlParameter -Command $cmd -Name '@OutputType'  -Type VarChar  -Size 20 -Value $OutputType
        Add-SqlParameter -Command $cmd -Name '@OutputText'  -Type NVarChar -Size -1 -Value $OutputText
        $null = $cmd.ExecuteNonQuery()
    }
    finally { $cmd.Dispose() }
}

# ======================================================================
# Abfrage des letzten Laufs (Abfragemodus)
# ======================================================================
function Get-LastExecution {
    # Liefert den zuletzt gestarteten Lauf eines Scripts (optional nur auf einem Rechner) oder $null.
    param([Parameter(Mandatory)][string]$Name, [string]$Computer)
    $computerValue = if ([string]::IsNullOrWhiteSpace($Computer)) { $null } else { $Computer }
    $cmd = New-SqlCommandObject -ProcedureName 'dbo.PS_GetLastExecution'
    try {
        Add-SqlParameter -Command $cmd -Name '@ScriptName'   -Type NVarChar -Size 200 -Value $Name
        Add-SqlParameter -Command $cmd -Name '@ComputerName' -Type NVarChar -Size 256 -Value $computerValue
        $reader = $cmd.ExecuteReader()
        try {
            if (-not $reader.Read()) { return $null }
            [pscustomobject]@{
                ExecutionID       = [guid]$reader['ExecutionID']
                ScriptID          = [int]$reader['ScriptID']
                ScriptName        = [string]$reader['ScriptName']
                ScriptVersion     = [int]$reader['ScriptVersion']
                ComputerName      = [string]$reader['ComputerName']
                RunsAs            = [string]$reader['RunsAs']
                Status            = [string]$reader['ExecutionStatus']
                StartedUtc        = [datetime]$reader['StartedAt']
                FinishedUtc       = $(if ($reader['FinishedAt'] -is [System.DBNull]) { $null } else { [datetime]$reader['FinishedAt'] })
                DurationMs        = $(if ($reader['DurationMs'] -is [System.DBNull]) { $null } else { [long]$reader['DurationMs'] })
                ErrorMessage      = $(if ($reader['ErrorMessage'] -is [System.DBNull]) { $null } else { [string]$reader['ErrorMessage'] })
                OutputCount       = [int]$reader['OutputCount']
            }
        }
        finally { $reader.Dispose() }
    }
    finally { $cmd.Dispose() }
}

function Get-ExecutionOutputRows {
    param([Parameter(Mandatory)][guid]$ExecutionId)
    $cmd  = New-SqlCommandObject -ProcedureName 'dbo.PS_GetExecutionOutput'
    $list = [System.Collections.Generic.List[object]]::new()
    try {
        Add-SqlParameter -Command $cmd -Name '@ExecutionID' -Type UniqueIdentifier -Value $ExecutionId
        $reader = $cmd.ExecuteReader()
        try {
            while ($reader.Read()) {
                $list.Add([pscustomobject]@{
                        OutputID   = [long]$reader['OutputID']
                        OutputType = [string]$reader['OutputType']
                        OutputText = $(if ($reader['OutputText'] -is [System.DBNull]) { '' } else { [string]$reader['OutputText'] })
                    })
            }
        }
        finally { $reader.Dispose() }
    }
    finally { $cmd.Dispose() }
    $list.ToArray()    # einzelne Objekte ausgeben (Pipeline entpackt das Array)
}

# ======================================================================
# Lokale Warteschlange (DPAPI-verschlüsselt, ACL-geschützt)
# ======================================================================
function Initialize-QueueDirectory {
    if (Test-Path -LiteralPath $QueuePath) { return }
    $null = New-Item -ItemType Directory -Path $QueuePath -Force
    try {
        $acl = Get-Acl -LiteralPath $QueuePath
        $acl.SetAccessRuleProtection($true, $false)    # Vererbung aus, geerbte Rechte entfernen
        $sids = @(
            [System.Security.Principal.WindowsIdentity]::GetCurrent().User,
            (New-Object System.Security.Principal.SecurityIdentifier 'S-1-5-18'),       # SYSTEM
            (New-Object System.Security.Principal.SecurityIdentifier 'S-1-5-32-544')    # Administratoren
        )
        foreach ($sid in $sids) {
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($sid, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
            $acl.AddAccessRule($rule)
        }
        Set-Acl -LiteralPath $QueuePath -AclObject $acl
    }
    catch {
        Write-ExecLog -Level WARNING -Message "Zugriffsrechte für '$QueuePath' konnten nicht eingeschränkt werden: $(Get-ErrorText $_)"
    }
}

function Write-QueueFile {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)]$Item)
    $json  = ConvertTo-Json -InputObject $Item -Depth 6 -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)

    if ($script:DpapiAvailable) {
        $payload = [System.Security.Cryptography.ProtectedData]::Protect($bytes, $script:DpapiEntropy, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        $marker  = [byte]0x44      # 'D' = DPAPI
    }
    else {
        $payload = $bytes
        $marker  = [byte]0x50      # 'P' = Klartext (nur ACL-Schutz)
    }
    $all = New-Object byte[] ($payload.Length + 1)
    $all[0] = $marker
    [System.Array]::Copy($payload, 0, $all, 1, $payload.Length)

    $tmp = "$Path.tmp"
    [System.IO.File]::WriteAllBytes($tmp, $all)
    Move-Item -LiteralPath $tmp -Destination $Path -Force
}

function Read-QueueFile {
    param([Parameter(Mandatory)][string]$Path)
    $all = [System.IO.File]::ReadAllBytes($Path)
    if ($all.Length -lt 2) { throw 'Warteschlangendatei ist leer oder unvollständig.' }
    $payload = New-Object byte[] ($all.Length - 1)
    [System.Array]::Copy($all, 1, $payload, 0, $payload.Length)

    switch ($all[0]) {
        0x44 {
            if (-not $script:DpapiAvailable) { throw 'DPAPI ist auf diesem System nicht verfügbar.' }
            $bytes = [System.Security.Cryptography.ProtectedData]::Unprotect($payload, $script:DpapiEntropy, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        }
        0x50 { $bytes = $payload }
        default { throw 'Unbekanntes Format der Warteschlangendatei.' }
    }
    [System.Text.Encoding]::UTF8.GetString($bytes) | ConvertFrom-Json
}

function Save-ToQueue {
    param([Parameter(Mandatory)]$Entry, [object[]]$Outputs)
    Initialize-QueueDirectory
    $item = [pscustomobject]@{
        FormatVersion = 1
        QueuedTicks   = [datetime]::UtcNow.Ticks
        Log           = $Entry
        Outputs       = @($Outputs)
    }
    $path = Join-Path $QueuePath ('{0}.psq' -f $Entry.ExecutionID)
    Write-QueueFile -Path $path -Item $item
    $path
}

function Send-QueuedExecutions {
    if (-not (Test-Path -LiteralPath $QueuePath)) { return }
    $files = @(Get-ChildItem -LiteralPath $QueuePath -Filter '*.psq' -File | Sort-Object -Property LastWriteTimeUtc)
    if ($files.Count -eq 0) { return }

    Write-ExecLog -Level INFO -Message "$($files.Count) zwischengespeicherte Ausführung(en) werden übertragen."
    foreach ($file in $files) {
        try { $item = Read-QueueFile -Path $file.FullName }
        catch {
            Write-ExecLog -Level ERROR -Message "Warteschlangendatei '$($file.Name)' ist nicht lesbar (beschädigt, manipuliert oder anderes Konto): $(Get-ErrorText $_). Datei wird als .corrupt abgelegt."
            try { Move-Item -LiteralPath $file.FullName -Destination ($file.FullName + '.corrupt') -Force } catch { }
            continue
        }

        $outputs = @($item.Outputs)
        $sent = 0
        try {
            Write-ExecutionLogToDb -Entry $item.Log
            foreach ($o in $outputs) {
                Write-ExecutionOutputToDb -ExecutionId $item.Log.ExecutionID -OutputType $o.OutputType -OutputText $o.OutputText
                $sent++
            }
            Remove-Item -LiteralPath $file.FullName -Force
            Write-ExecLog -Level INFO -Message "Nachgesendet: ExecutionID $($item.Log.ExecutionID) ($sent Ausgabeeinträge)."
        }
        catch {
            Write-ExecLog -Level WARNING -Message "Nachsenden von '$($file.Name)' fehlgeschlagen, bleibt in der Warteschlange: $(Get-ErrorText $_)"
            if ($sent -gt 0) {
                # Bereits übertragene Ausgaben nicht erneut senden.
                try {
                    $item.Outputs = @($outputs | Select-Object -Skip $sent)
                    Write-QueueFile -Path $file.FullName -Item $item
                }
                catch { Write-ExecLog -Level WARNING -Message "Fortschritt konnte nicht gespeichert werden: $(Get-ErrorText $_)" }
            }
            break    # Verbindung vermutlich weg, weitere Dateien nicht versuchen
        }
    }
}

# ======================================================================
# Protokoll-Eintrag und Übertragung
# ======================================================================
function New-LogEntry {
    param(
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][datetime]$StartedAt,
        $FinishedAt,
        $DurationMs,
        $ErrorMessage
    )
    $c = $script:Ctx
    [pscustomobject]@{
        ExecutionID       = $script:ExecutionId.ToString()
        ScriptID          = $c.ScriptID
        ScriptVersion     = $c.ScriptVersion
        ComputerName      = $c.ComputerName
        RunsAs            = $c.RunsAs
        PowerShellVersion = $c.PowerShellVersion
        PowerShellEdition = $c.PowerShellEdition
        OSVersion         = $c.OSVersion
        InstalledModules  = $c.InstalledModules
        ExecutionStatus   = $Status
        StartedTicks      = $StartedAt.Ticks
        FinishedTicks     = $(if ($null -ne $FinishedAt) { ([datetime]$FinishedAt).Ticks } else { $null })
        DurationMs        = $DurationMs
        ErrorMessage      = $ErrorMessage
    }
}

function Send-ExecutionResult {
    # Überträgt Protokoll und Ausgaben. Bei Fehlern wird alles Ungesendete lokal zwischengespeichert.
    # Rückgabe: 'Sent', 'Queued' oder 'Lost'
    param([Parameter(Mandatory)]$Entry, [object[]]$Outputs)
    $sent = 0
    try {
        Open-SqlConnection     # frische Verbindung (lange Läufe können die alte verloren haben)
        Write-ExecutionLogToDb -Entry $Entry
        foreach ($o in $Outputs) {
            Write-ExecutionOutputToDb -ExecutionId $Entry.ExecutionID -OutputType $o.OutputType -OutputText $o.OutputText
            $sent++
        }
        return 'Sent'
    }
    catch {
        Write-ExecLog -Level WARNING -Message "Ergebnis konnte nicht an SQL Server übertragen werden: $(Get-ErrorText $_)"
        try {
            $remaining = @($Outputs | Select-Object -Skip $sent)
            $path = Save-ToQueue -Entry $Entry -Outputs $remaining
            Write-ExecLog -Level WARNING -Message "Ergebnis lokal zwischengespeichert: $path"
            return 'Queued'
        }
        catch {
            Write-ExecLog -Level ERROR -Message "Ergebnis konnte auch nicht lokal gespeichert werden (Daten gehen verloren): $(Get-ErrorText $_)"
            return 'Lost'
        }
    }
}

# ======================================================================
# Script-Ausführung (eigener Runspace)
# ======================================================================
function Add-OutputEntry {
    param([string]$Type, [string]$Text, [string]$Prefix = '')
    if ([string]::IsNullOrWhiteSpace($Text)) { return }
    $run = $script:Run
    if ($run.Entries.Count -ge $MaxOutputEntries) { $run.Dropped++; return }
    if ($Prefix) {
        # Bei mehreren Scripten: Scriptname voranstellen (bei OUTPUT als eigene Zeile, damit Tabellen intakt bleiben)
        if ($Type -eq 'OUTPUT') { $Text = $Prefix + "`n" + $Text } else { $Text = "$Prefix $Text" }
    }
    if ($Text.Length -gt $MaxOutputChars) {
        $orig = $Text.Length
        $Text = $Text.Substring(0, $MaxOutputChars) + "`n[... gekürzt, ursprüngliche Länge: $orig Zeichen]"
    }
    $run.Entries.Add([pscustomobject]@{ OutputType = $Type; OutputText = $Text })
}

function Add-StreamEntries {
    # Überträgt Ausgabe und Streams eines Scripts in die Liste der Ausgabeeinträge.
    param($Ps, $OutColl, [string]$Terminating, [string]$Prefix)
    if ($OutColl.Count -gt 0) {
        try { $text = ($OutColl | Out-String -Width 200).TrimEnd() }
        catch { $text = (($OutColl | ForEach-Object { [string]$_ }) -join [Environment]::NewLine) }
        Add-OutputEntry -Type 'OUTPUT' -Text $text -Prefix $Prefix
    }
    if ($null -ne $Ps) {
        foreach ($e in $Ps.Streams.Error)       { Add-OutputEntry -Type 'ERROR'       -Text (Format-ErrorRecord -Record $e) -Prefix $Prefix }
        if (-not [string]::IsNullOrEmpty($Terminating)) { Add-OutputEntry -Type 'ERROR' -Text "Terminierender Fehler: $Terminating" -Prefix $Prefix }
        foreach ($w in $Ps.Streams.Warning)     { Add-OutputEntry -Type 'WARNING'     -Text $w.Message -Prefix $Prefix }
        foreach ($i in $Ps.Streams.Information) { Add-OutputEntry -Type 'INFORMATION' -Text $i.ToString() -Prefix $Prefix }
        foreach ($v in $Ps.Streams.Verbose)     { Add-OutputEntry -Type 'VERBOSE'     -Text $v.Message -Prefix $Prefix }
        foreach ($d in $Ps.Streams.Debug)       { Add-OutputEntry -Type 'DEBUG'       -Text $d.Message -Prefix $Prefix }
    }
}

function Invoke-OneScript {
    # Führt ein Script in einem bereits geöffneten Runspace aus. Der Runspace bleibt danach bestehen,
    # sodass Funktionen und Variablen für folgende Scripte verfügbar sind.
    param(
        [Parameter(Mandatory)]$Runspace,
        [Parameter(Mandatory)]$Item,
        [Parameter(Mandatory)][System.Diagnostics.Stopwatch]$Total,
        [bool]$Multi
    )

    $res = [pscustomobject]@{ Status = 'SUCCESS'; Message = $null; Stop = $false; TimedOut = $false }
    $prefix = if ($Multi) { "[$($Item.Name)]" } else { '' }
    $label  = if ($Multi) { "Script '$($Item.Name)': " } else { '' }
    $where  = if ($Multi) { " bei '$($Item.Name)'" } else { '' }

    $ps = $null
    $outColl = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
    $terminating = $null
    try {
        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.Runspace = $Runspace
        $null = $ps.AddScript([string]$Item.Source)
        if ($null -ne $Item.Parameters -and $Item.Parameters.Count -gt 0) { $null = $ps.AddParameters($Item.Parameters) }

        $inColl = New-Object 'System.Management.Automation.PSDataCollection[psobject]'
        $inColl.Complete()
        $async = $ps.BeginInvoke($inColl, $outColl)

        while (-not $async.AsyncWaitHandle.WaitOne(500)) {
            if ($TimeoutSec -gt 0 -and $Total.Elapsed.TotalSeconds -ge $TimeoutSec) {
                $res.TimedOut = $true
                Write-ExecLog -Level WARNING -Message "Zeitüberschreitung nach $TimeoutSec s$where, das Script wird gestoppt."
                $ps.Stop()
                break
            }
        }

        try { $null = $ps.EndInvoke($async) }
        catch {
            if (-not $res.TimedOut) {
                $base = $_.Exception.GetBaseException()
                if ($base -is [System.Management.Automation.RuntimeException] -and $null -ne $base.ErrorRecord) {
                    $terminating = Format-ErrorRecord -Record $base.ErrorRecord
                }
                else { $terminating = $base.Message }
            }
        }

        # Status dieses Scripts
        $errCount  = $ps.Streams.Error.Count
        $warnCount = $ps.Streams.Warning.Count
        if ($res.TimedOut) {
            $res.Status  = 'CANCELLED'
            $res.Message = "Zeitüberschreitung${where}: Die Ausführung wurde nach $TimeoutSec Sekunden gestoppt."
            $res.Stop    = $true
        }
        elseif ($null -ne $terminating) {
            $res.Status  = 'ERROR'
            $res.Message = $label + $terminating
            $res.Stop    = $true
        }
        elseif ($errCount -gt 0) {
            $res.Status  = 'ERROR'
            $res.Message = "$label$errCount Fehler im Script. Erster Fehler: $($ps.Streams.Error[0].ToString())"
        }
        elseif ($warnCount -gt 0) {
            $res.Status = 'WARNING'
        }
    }
    catch {
        if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
        $terminating = "Interner Fehler bei der Ausführung des Scripts: $(Get-ErrorText $_)"
        $res.Status  = 'ERROR'
        $res.Message = $label + $terminating
        $res.Stop    = $true
    }
    finally {
        # Bei Abbruch (Strg+C) laufendes Script stoppen und trotzdem Ausgaben einsammeln.
        if ($null -ne $ps) {
            try { if ($ps.InvocationStateInfo.State -eq 'Running') { $ps.Stop() } } catch { }
        }
        try { Add-StreamEntries -Ps $ps -OutColl $outColl -Terminating $terminating -Prefix $prefix }
        catch { Write-ExecLog -Level WARNING -Message "Ausgaben konnten nicht vollständig erfasst werden: $(Get-ErrorText $_)" }
        if ($null -ne $ps) { $ps.Dispose() }
    }

    $res
}

function Invoke-ScriptSession {
    # Führt ein oder mehrere Scripte nacheinander im selben Runspace aus.
    # Das Ergebnis wird in $script:Run geschrieben (auch bei einem Abbruch verfügbar).
    param([Parameter(Mandatory)][object[]]$Items)

    $multi   = $Items.Count -gt 1
    $rs      = $null
    $total   = [System.Diagnostics.Stopwatch]::StartNew()
    $rank    = @{ SUCCESS = 0; WARNING = 1; ERROR = 2; CANCELLED = 3 }
    $overall = 'SUCCESS'
    $overallMessage = $null
    try {
        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.Open()
        $rs.SessionStateProxy.SetVariable('ProgressPreference', 'SilentlyContinue')
        $rs.SessionStateProxy.SetVariable('PS_ExecutionID', $script:ExecutionId.ToString())
        if ($CaptureVerbose) { $rs.SessionStateProxy.SetVariable('VerbosePreference', 'Continue') }
        if ($CaptureDebug)   { $rs.SessionStateProxy.SetVariable('DebugPreference', 'Continue') }

        if ($multi) {
            $list = ($Items | ForEach-Object { '{0} (Version {1})' -f $_.Name, $_.Version }) -join ', '
            Add-OutputEntry -Type 'INFORMATION' -Text ("Session mit {0} Scripten im selben Runspace: {1}. Hauptscript: '{2}'." -f $Items.Count, $list, $Items[$Items.Count - 1].Name)
        }

        for ($i = 0; $i -lt $Items.Count; $i++) {
            $item = $Items[$i]
            if ($multi) { Write-ExecLog -Level INFO -Message "Script $($i + 1) von $($Items.Count) gestartet: '$($item.Name)'." }

            $r = Invoke-OneScript -Runspace $rs -Item $item -Total $total -Multi $multi

            if ($rank[$r.Status] -gt $rank[$overall]) { $overall = $r.Status }
            if ($r.Status -eq 'CANCELLED') { $overallMessage = $r.Message }
            elseif ($r.Status -eq 'ERROR' -and $null -eq $overallMessage) { $overallMessage = $r.Message }

            if ($r.Stop) {
                for ($j = $i + 1; $j -lt $Items.Count; $j++) {
                    Add-OutputEntry -Type 'WARNING' -Text ("Script '{0}' wurde nicht ausgeführt, weil '{1}' vorzeitig beendet wurde." -f $Items[$j].Name, $item.Name)
                }
                break
            }
        }

        $script:Run.Status       = $overall
        $script:Run.ErrorMessage = $overallMessage
    }
    catch {
        if ($_.Exception -is [System.Management.Automation.PipelineStoppedException]) { throw }
        $msg = "Interner Fehler bei der Ausführung der Scripte: $(Get-ErrorText $_)"
        Add-OutputEntry -Type 'ERROR' -Text $msg
        $script:Run.Status       = 'ERROR'
        $script:Run.ErrorMessage = $msg
    }
    finally {
        if ($script:Run.Dropped -gt 0) {
            $script:Run.Entries.Add([pscustomobject]@{
                    OutputType = 'WARNING'
                    OutputText = "Es wurden $($script:Run.Dropped) weitere Ausgabeeinträge verworfen (Limit: $MaxOutputEntries)."
                })
        }
        if ($null -ne $rs) { try { $rs.Close() } catch { }; $rs.Dispose() }
    }
}

function Get-ScriptNameList {
    # -ScriptName akzeptiert mehrere Namen. Bei "-File" (Aufgabenplanung) kommen sie als ein Text "A,B" an.
    @($ScriptName | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

# ======================================================================
# Hauptablauf
# ======================================================================
function Invoke-Main {
    # --- Verbindung ---
    try {
        Open-SqlConnection
        Write-ExecLog -Level INFO -Message ("Verbunden mit '{0}' / '{1}' (verschlüsselt, {2})." -f $SqlServer, $Database, $(if ($ValidateServerCertificate) { 'Zertifikatsprüfung aktiv' } else { 'Zertifikatsprüfung deaktiviert' }))
    }
    catch {
        Write-ExecLog -Level ERROR -Message "Verbindung zu '$SqlServer' / '$Database' fehlgeschlagen: $(Get-ErrorText $_)"
        $script:ExitCode = 2
        return
    }

    try {
        # --- Warteschlange nachsenden ---
        if ($script:SetName -eq 'Flush' -or -not $NoFlush) { Send-QueuedExecutions }
        if ($script:SetName -eq 'Flush') { return }

        # --- Auflisten ---
        if ($script:SetName -eq 'List') {
            Get-RepositoryScripts | Select-Object -Property ScriptID, ScriptName, ScriptVersion, Description, ModifiedAt
            return
        }

        # --- Abfragemodus: letzten Lauf anzeigen (-ShowLastRun / -ShowOutput) ---
        if ($script:SetName -eq 'ShowLastRun' -or $script:SetName -eq 'ShowOutput') {
            $showNames = @(Get-ScriptNameList)
            if ($showNames.Count -ne 1) {
                Write-ExecLog -Level ERROR -Message 'Der Abfragemodus (-ShowOutput / -ShowLastRun) unterstützt genau einen Scriptnamen. Bei einer Session mit mehreren Scripten den Namen des Hauptscripts (des letzten) angeben.'
                $script:ExitCode = 3
                return
            }
            $showName = $showNames[0]
            try { $last = Get-LastExecution -Name $showName -Computer $ComputerName }
            catch {
                Write-ExecLog -Level ERROR -Message "Abfrage des letzten Laufs fehlgeschlagen (PS_ViewerRole erforderlich?): $(Get-ErrorText $_)"
                $script:ExitCode = 3
                return
            }
            if (-not $last) {
                $where = if ($ComputerName) { " auf '$ComputerName'" } else { '' }
                Write-ExecLog -Level WARNING -Message "Für '$showName'$where wurde kein Lauf gefunden."
                $script:ExitCode = 3
                return
            }

            $started  = [datetime]::SpecifyKind($last.StartedUtc, [System.DateTimeKind]::Utc).ToLocalTime()
            $finished = if ($null -ne $last.FinishedUtc) { [datetime]::SpecifyKind($last.FinishedUtc, [System.DateTimeKind]::Utc).ToLocalTime() } else { $null }
            $duration = if ($null -ne $last.DurationMs) { [timespan]::FromMilliseconds([double]$last.DurationMs) } else { $null }

            if ($script:SetName -eq 'ShowLastRun') {
                [pscustomobject]@{
                    ScriptName    = $last.ScriptName
                    ScriptVersion = $last.ScriptVersion
                    Status        = $last.Status
                    Started       = $started
                    Finished      = $finished
                    Duration      = $duration
                    ComputerName  = $last.ComputerName
                    RunsAs        = $last.RunsAs
                    OutputEntries = $last.OutputCount
                    ExecutionID   = $last.ExecutionID
                    ErrorMessage  = $last.ErrorMessage
                }
                return
            }

            # -ShowOutput: gespeicherte Ausgaben des letzten Laufs
            Write-Host ("Letzter Lauf von '{0}' (Version {1}) auf {2} als {3}" -f $last.ScriptName, $last.ScriptVersion, $last.ComputerName, $last.RunsAs) -ForegroundColor Cyan
            Write-Host ("Status: {0}   Start: {1:yyyy-MM-dd HH:mm:ss}   Dauer: {2}   ExecutionID: {3}" -f `
                    $last.Status, $started, $(if ($null -ne $duration) { $duration.ToString() } else { '-' }), $last.ExecutionID) -ForegroundColor Cyan
            if ($last.Status -eq 'STARTED') {
                Write-Host 'Der Lauf ist noch aktiv oder wurde beendet, ohne das Ende zu melden (Status STARTED).' -ForegroundColor Yellow
            }
            if (-not [string]::IsNullOrWhiteSpace($last.ErrorMessage)) {
                Write-Host "Fehlermeldung: $($last.ErrorMessage)" -ForegroundColor Red
            }
            Write-Host ''

            $rows = @(Get-ExecutionOutputRows -ExecutionId $last.ExecutionID)
            if ($OutputType) { $rows = @($rows | Where-Object { $OutputType -contains $_.OutputType }) }

            if ($rows.Count -eq 0) {
                Write-Host '(keine Ausgabeeinträge gespeichert' -NoNewline
                if ($OutputType) { Write-Host " für Typ $($OutputType -join ', ')" -NoNewline }
                Write-Host ')'
                return
            }

            $colors = @{ ERROR = 'Red'; WARNING = 'Yellow'; INFORMATION = 'Cyan'; VERBOSE = 'DarkGray'; DEBUG = 'DarkGray' }
            foreach ($r in $rows) {
                $line  = '[{0}] {1}' -f $r.OutputType, $r.OutputText
                $color = $colors[$r.OutputType]
                if ($color) { Write-Host $line -ForegroundColor $color } else { Write-Host $line }
            }
            return
        }

        # --- Scripte ermitteln (eines oder mehrere; das letzte ist das Hauptscript) ---
        $ids = [System.Collections.Generic.List[int]]::new()
        if ($script:SetName -eq 'ByName') {
            $available = @(Get-RepositoryScripts)
            $missing = [System.Collections.Generic.List[string]]::new()
            foreach ($n in @(Get-ScriptNameList)) {
                $found = $available | Where-Object { $_.ScriptName -eq $n } | Select-Object -First 1
                if (-not $found) { $missing.Add($n) }
                elseif ($ids.Contains($found.ScriptID)) {
                    Write-ExecLog -Level WARNING -Message "Script '$n' wurde mehrfach angegeben und wird nur einmal ausgeführt."
                }
                else { $ids.Add($found.ScriptID) }
            }
            if ($missing.Count -gt 0) {
                Write-ExecLog -Level ERROR -Message ("Script(e) nicht gefunden oder nicht aktiv: {0}" -f (($missing | ForEach-Object { "'$_'" }) -join ', '))
                $script:ExitCode = 3
                return
            }
        }
        else { $ids.Add($ScriptID) }

        # --- Scripte abrufen (alle, bevor etwas ausgeführt wird) ---
        $recs = [System.Collections.Generic.List[object]]::new()
        foreach ($id in $ids) {
            try { $one = Get-RepositoryScript -Id $id }
            catch {
                Write-ExecLog -Level ERROR -Message "Script (ID $id) konnte nicht abgerufen werden: $(Get-ErrorText $_)"
                $script:ExitCode = 3
                return
            }
            $recs.Add($one)
            Write-ExecLog -Level INFO -Message "Script abgerufen: '$($one.ScriptName)' (ID $($one.ScriptID)), Version $($one.ScriptVersion)."
        }
        $rec   = $recs[$recs.Count - 1]      # Hauptscript = das letzte
        $multi = $recs.Count -gt 1
        $names = ($recs | ForEach-Object { $_.ScriptName }) -join ' -> '

        # --- Systeminformationen und Kontext (Protokoll läuft unter dem Hauptscript) ---
        $sys = Get-SystemInfo
        $script:Ctx = [pscustomobject]@{
            ScriptID          = $rec.ScriptID
            ScriptName        = $rec.ScriptName
            ScriptVersion     = $rec.ScriptVersion
            ComputerName      = $sys.ComputerName
            RunsAs            = $sys.RunsAs
            PowerShellVersion = $sys.PowerShellVersion
            PowerShellEdition = $sys.PowerShellEdition
            OSVersion         = $sys.OSVersion
            InstalledModules  = $sys.InstalledModules
        }
        $script:ExecutionId = [guid]::NewGuid()

        # --- Integritätsprüfung aller Scripte ---
        foreach ($s in $recs) {
            $localHash = Get-SourceHash -Source $s.ScriptSource
            if ($localHash -ne $s.HashSHA256) {
                $msg = "Hash-Abweichung bei '$($s.ScriptName)': berechnet=$localHash, Datenbank=$($s.HashSHA256). Es wird kein Script ausgeführt."
                Write-ExecLog -Level ERROR -Message $msg
                $now = [datetime]::UtcNow
                $entry = New-LogEntry -Status 'ERROR' -StartedAt $now -FinishedAt $now -DurationMs 0 -ErrorMessage $msg
                $null = Send-ExecutionResult -Entry $entry -Outputs @()
                $script:ExitCode = 4
                return
            }
            Write-ExecLog -Level DEBUG -Message "Hash geprüft: '$($s.ScriptName)' $localHash"
        }

        # --- Ausführung ---
        $script:Run = [pscustomobject]@{
            Status       = 'CANCELLED'
            ErrorMessage = 'Die Ausführung wurde vorzeitig abgebrochen.'
            Entries      = [System.Collections.Generic.List[object]]::new()
            Dropped      = 0
        }

        $items = [System.Collections.Generic.List[object]]::new()
        for ($k = 0; $k -lt $recs.Count; $k++) {
            $isMain = ($k -eq $recs.Count - 1)
            $items.Add([pscustomobject]@{
                    Name       = $recs[$k].ScriptName
                    Version    = $recs[$k].ScriptVersion
                    Source     = $recs[$k].ScriptSource
                    Parameters = $(if ($isMain) { $ScriptParameters } else { $null })
                })
        }

        $startedAt = [datetime]::UtcNow
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        try {
            $startEntry = New-LogEntry -Status 'STARTED' -StartedAt $startedAt -FinishedAt $null -DurationMs $null -ErrorMessage $null
            Write-ExecutionLogToDb -Entry $startEntry
        }
        catch {
            Write-ExecLog -Level WARNING -Message "Startmeldung konnte nicht gespeichert werden (das Endergebnis wird nachgetragen): $(Get-ErrorText $_)"
        }

        if ($multi) {
            Write-ExecLog -Level INFO -Message "Ausführung gestartet: Session mit $($recs.Count) Scripten ($names), Hauptscript '$($rec.ScriptName)' Version $($rec.ScriptVersion) auf $($sys.ComputerName) als $($sys.RunsAs)."
        }
        else {
            Write-ExecLog -Level INFO -Message "Ausführung gestartet: '$($rec.ScriptName)' Version $($rec.ScriptVersion) auf $($sys.ComputerName) als $($sys.RunsAs)."
        }

        try {
            Invoke-ScriptSession -Items $items.ToArray()
        }
        finally {
            $sw.Stop()
            $finishedAt = [datetime]::UtcNow
            $entry = New-LogEntry -Status $script:Run.Status -StartedAt $startedAt -FinishedAt $finishedAt `
                -DurationMs ([long]$sw.ElapsedMilliseconds) -ErrorMessage $script:Run.ErrorMessage
            $delivery = Send-ExecutionResult -Entry $entry -Outputs @($script:Run.Entries)
        }

        Write-ExecLog -Level INFO -Message ("Ausführung beendet: Status {0}, Dauer {1} ms, {2} Ausgabeeinträge, Übertragung: {3}." -f `
                $script:Run.Status, [long]$sw.ElapsedMilliseconds, $script:Run.Entries.Count, $delivery)

        if ($PrintOutput) {
            foreach ($e in $script:Run.Entries) {
                Write-Host ("[{0}] {1}" -f $e.OutputType, $e.OutputText)
            }
        }

        $script:ExitCode = if ($script:Run.Status -in 'SUCCESS', 'WARNING') { 0 } else { 1 }

        [pscustomobject]@{
            ExecutionID = $script:ExecutionId
            ScriptName  = $rec.ScriptName
            ScriptID    = $rec.ScriptID
            Version     = $rec.ScriptVersion
            Scripts     = $names
            Status      = $script:Run.Status
            DurationMs  = [long]$sw.ElapsedMilliseconds
            Delivery    = $delivery
            Message     = $script:Run.ErrorMessage
        }
    }
    finally {
        if ($null -ne $script:Conn) { try { $script:Conn.Dispose() } catch { } }
    }
}

# ======================================================================
# Initialisierung
# ======================================================================
try {
    if (-not (Test-Path -LiteralPath $LogPath)) { $null = New-Item -ItemType Directory -Path $LogPath -Force }
    $script:LogFile = Join-Path $LogPath ('PS_Execute_{0:yyyyMMdd}.log' -f (Get-Date))
}
catch {
    $script:LogFile = $null
    Write-Warning "Logverzeichnis '$LogPath' nicht verfügbar: $($_.Exception.Message). Es wird nur auf der Konsole protokolliert."
}

# SQL-Client ermitteln (Windows PowerShell 5.1: System.Data explizit laden)
if ($PSVersionTable.PSVersion.Major -lt 6) { Add-Type -AssemblyName System.Data }
foreach ($ns in 'System.Data.SqlClient', 'Microsoft.Data.SqlClient') {
    if ("$ns.SqlConnection" -as [type]) { $script:Provider = $ns; break }
}
if (-not $script:Provider) {
    Write-ExecLog -Level ERROR -Message 'Kein SQL-Client gefunden (System.Data.SqlClient bzw. Microsoft.Data.SqlClient).'
    exit 2
}

# DPAPI für die Warteschlange
foreach ($asm in 'System.Security', 'System.Security.Cryptography.ProtectedData') {
    try { Add-Type -AssemblyName $asm -ErrorAction Stop } catch { }
}
if ('System.Security.Cryptography.ProtectedData' -as [type]) { $script:DpapiAvailable = $true }
else { Write-ExecLog -Level WARNING -Message 'DPAPI nicht verfügbar: Die lokale Warteschlange wird nur über Dateirechte geschützt.' }

Invoke-Main
exit $script:ExitCode