#Requires -Version 5.1
<#
.SYNOPSIS
    Lädt PowerShell-Scripte (.ps1) über die Stored Procedure dbo.PS_UploadScript
    in das zentrale Script-Repository (SQL Server) hoch.

.DESCRIPTION
    - Der Scriptname wird automatisch aus dem Dateinamen (ohne Endung .ps1) übernommen.
    - Legt ein Script neu an oder aktualisiert ein vorhandenes Script (Zuordnung über den Scriptnamen).
    - Die Version wird ausschließlich serverseitig vergeben (PS_UploadScript).
    - Berechnet lokal den SHA-256-Hash (UTF-16LE, identisch zur Berechnung in SQL Server)
      und vergleicht ihn nach dem Upload mit dem vom Server gespeicherten Hash.
    - Prüft vor dem Upload die PowerShell-Syntax (kann übersprungen werden).
    - Überspringt unveränderte Scripte, damit keine unnötigen Versionen entstehen (außer mit -Force).
    - Protokolliert Upload und Ergebnis in einer lokalen Logdatei (kein Sourcecode im Log).
    - Verwendet Windows Integrated Authentication.

.PARAMETER FilePath
    Pfad zur .ps1-Datei (Wildcards und Pipeline möglich).

.PARAMETER Description
    Beschreibung (max. 1000 Zeichen). Wird sie bei einem Update nicht angegeben,
    bleibt die bisherige Beschreibung erhalten.

.PARAMETER SqlServer
    SQL-Server-Instanz, z. B. "SQL01" oder "SQL01\INST1".

.PARAMETER Database
    Datenbankname. Standard: PowerShellAutomation.

.PARAMETER Inactive
    Script inaktiv speichern (wird dann nicht mehr über PS_GetScripts angeboten).

.PARAMETER Force
    Upload auch dann durchführen, wenn sich Inhalt und Beschreibung nicht geändert haben.

.PARAMETER SkipSyntaxCheck
    Lokale Syntaxprüfung überspringen.

.PARAMETER ValidateServerCertificate
    Standardmäßig wird das SQL-Server-Zertifikat NICHT geprüft (Betrieb im internen,
    geschützten Netz). Mit diesem Schalter wird die Zertifikatsprüfung aktiviert.
    Die Verbindung bleibt in beiden Fällen verschlüsselt.

.PARAMETER LogPath
    Verzeichnis für die lokale Logdatei. Standard: %LOCALAPPDATA%\PSAutomation\Logs

.EXAMPLE
    .\PS_Upload.ps1 -FilePath .\Get-DiskReport.ps1 -SqlServer SQL01 -Description 'Festplattenbericht'

.EXAMPLE
    Get-ChildItem .\Scripts\*.ps1 | .\PS_Upload.ps1 -SqlServer SQL01 -WhatIf

.NOTES
    Exitcode 0 = alle Uploads erfolgreich (oder unverändert), 1 = mindestens ein Fehler, 2 = Verbindungsfehler.
    Lauffähig unter Windows PowerShell 5.1 und PowerShell 7+.
    Benötigt EXECUTE-Rechte auf dbo.PS_GetScripts und dbo.PS_UploadScript (PS_UploaderRole).
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('Path', 'FullName')]
    [string[]]$FilePath,

    [ValidateLength(0, 1000)]
    [string]$Description,

    [string]$SqlServer = '#DBSERVER#',

    [string]$Database = '#DBNAME#',

    [switch]$Inactive,
    [switch]$Force,
    [switch]$SkipSyntaxCheck,
    [switch]$ValidateServerCertificate,

    [ValidateRange(5, 600)]
    [int]$CommandTimeoutSec = 60,

    [string]$LogPath = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'PSAutomation\Logs')
)

begin {
    Set-StrictMode -Version 3.0
    $ErrorActionPreference = 'Stop'

    $script:HadError  = $false
    $script:Conn      = $null
    $script:Provider  = $null
    $script:Counters  = @{ Created = 0; Updated = 0; Unchanged = 0; Failed = 0; WhatIf = 0 }

    # ---------- Logging ----------
    try {
        if (-not (Test-Path -LiteralPath $LogPath)) {
            $null = New-Item -ItemType Directory -Path $LogPath -Force
        }
        $script:LogFile = Join-Path $LogPath ('PS_Upload_{0:yyyyMMdd}.log' -f (Get-Date))
    }
    catch {
        $script:LogFile = $null
        Write-Warning "Logverzeichnis '$LogPath' nicht verfügbar: $($_.Exception.Message). Es wird nur auf der Konsole protokolliert."
    }

    function Write-UploadLog {
        param(
            [ValidateSet('INFO', 'WARNING', 'ERROR', 'DEBUG')]
            [string]$Level = 'INFO',
            [Parameter(Mandatory)]
            [string]$Message
        )
        $line = '{0:yyyy-MM-dd HH:mm:ss.fff}Z [{1}] [{2}\{3}@{4}] {5}' -f `
            [datetime]::UtcNow, $Level, [Environment]::UserDomainName, [Environment]::UserName,
            [Environment]::MachineName, $Message

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

    # ---------- Hilfsfunktionen ----------
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

    function New-SqlCommandObject {
        param([Parameter(Mandatory)][string]$ProcedureName)
        $cmd = New-Object "$($script:Provider).SqlCommand" $ProcedureName, $script:Conn
        $cmd.CommandType    = [System.Data.CommandType]::StoredProcedure
        $cmd.CommandTimeout = $CommandTimeoutSec
        $cmd
    }

    function Get-RepositoryScripts {
        $cmd  = New-SqlCommandObject -ProcedureName 'dbo.PS_GetScripts'
        $list = [System.Collections.Generic.List[object]]::new()
        try {
            $reader = $cmd.ExecuteReader()
            try {
                while ($reader.Read()) {
                    $list.Add([pscustomobject]@{
                            ScriptID      = [int]$reader['ScriptID']
                            ScriptName    = [string]$reader['ScriptName']
                            Description   = if ($reader['Description'] -is [System.DBNull]) { $null } else { [string]$reader['Description'] }
                            ScriptVersion = [int]$reader['ScriptVersion']
                            HashSHA256    = [string]$reader['HashSHA256']
                        })
                }
            }
            finally { $reader.Dispose() }
        }
        finally { $cmd.Dispose() }
        # Einzelne Objekte ausgeben (kein vorangestelltes Komma): Die Pipeline entpackt das Array,
    # bei leerem Repository wird dadurch nichts ausgegeben.
    $list.ToArray()
    }

    function Invoke-UploadProcedure {
        param(
            [Nullable[int]]$ScriptID,
            [Parameter(Mandatory)][string]$Name,
            [string]$Desc,
            [Parameter(Mandatory)][string]$Source,
            [Parameter(Mandatory)][bool]$Active
        )
        $cmd = New-SqlCommandObject -ProcedureName 'dbo.PS_UploadScript'
        try {
            $p = $cmd.Parameters.Add('@ScriptID', [System.Data.SqlDbType]::Int)
            $p.Value = if ($null -eq $ScriptID) { [System.DBNull]::Value } else { [int]$ScriptID }

            $p = $cmd.Parameters.Add('@ScriptName', [System.Data.SqlDbType]::NVarChar, 200)
            $p.Value = $Name

            $p = $cmd.Parameters.Add('@Description', [System.Data.SqlDbType]::NVarChar, 1000)
            $p.Value = if ([string]::IsNullOrEmpty($Desc)) { [System.DBNull]::Value } else { $Desc }

            $p = $cmd.Parameters.Add('@ScriptSource', [System.Data.SqlDbType]::NVarChar, -1)
            $p.Value = $Source

            $p = $cmd.Parameters.Add('@IsActive', [System.Data.SqlDbType]::Bit)
            $p.Value = $Active

            $reader = $cmd.ExecuteReader()
            try {
                if (-not $reader.Read()) { throw 'PS_UploadScript hat kein Ergebnis geliefert.' }
                [pscustomobject]@{
                    ScriptID      = [int]$reader['ScriptID']
                    ScriptName    = [string]$reader['ScriptName']
                    ScriptVersion = [int]$reader['ScriptVersion']
                    HashSHA256    = [string]$reader['HashSHA256']
                }
            }
            finally { $reader.Dispose() }
        }
        finally { $cmd.Dispose() }
    }

    # ---------- SQL-Provider ermitteln ----------
    # Windows PowerShell 5.1: System.Data (enthält System.Data.SqlClient) explizit laden.
    if ($PSVersionTable.PSVersion.Major -lt 6) {
        Add-Type -AssemblyName System.Data
    }
    foreach ($ns in 'System.Data.SqlClient', 'Microsoft.Data.SqlClient') {
        if ("$ns.SqlConnection" -as [type]) { $script:Provider = $ns; break }
    }
    if (-not $script:Provider) {
        Write-UploadLog -Level ERROR -Message 'Kein SQL-Client gefunden (System.Data.SqlClient bzw. Microsoft.Data.SqlClient).'
        exit 2
    }

    # ---------- Verbindung öffnen ----------
    $connString = 'Server={0};Database={1};Integrated Security=True;Encrypt=True;TrustServerCertificate={2};Application Name=PS_Upload.ps1;Connect Timeout=15' -f `
        $SqlServer, $Database, (-not $ValidateServerCertificate.IsPresent)

    try {
        $script:Conn = New-Object "$($script:Provider).SqlConnection" $connString
        $script:Conn.Open()
        $certInfo = if ($ValidateServerCertificate) { 'Zertifikatsprüfung aktiv' } else { 'Zertifikatsprüfung deaktiviert' }
        Write-UploadLog -Message "Verbunden mit '$SqlServer' / '$Database' (verschlüsselt, $certInfo)."
    }
    catch {
        Write-UploadLog -Level ERROR -Message "Verbindung zu '$SqlServer' / '$Database' fehlgeschlagen: $(Get-ErrorText $_)"
        exit 2
    }
}

process {
    # Wildcards auflösen
    $files = foreach ($path in $FilePath) {
        try { Resolve-Path -Path $path | Select-Object -ExpandProperty ProviderPath }
        catch {
            Write-UploadLog -Level ERROR -Message "Datei nicht gefunden: '$path'"
            $script:HadError = $true
            $script:Counters.Failed++
        }
    }

    foreach ($file in $files) {
        $result = [pscustomobject]@{
            File       = $file
            ScriptName = $null
            ScriptID   = $null
            Version    = $null
            Action     = 'None'
            Hash       = $null
            Status     = 'Failed'
            Message    = ''
        }

        try {
            if ([System.IO.Path]::GetExtension($file) -ne '.ps1') {
                throw 'Nur .ps1-Dateien werden unterstützt.'
            }

            # Datei lesen (BOM-Erkennung für UTF-8/UTF-16)
            $source = [System.IO.File]::ReadAllText($file, [System.Text.UTF8Encoding]::new($false))
            if ([string]::IsNullOrWhiteSpace($source)) { throw 'Die Datei ist leer.' }

            # Syntaxprüfung
            if (-not $SkipSyntaxCheck) {
                $tokens = $null; $parseErrors = $null
                $null = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$tokens, [ref]$parseErrors)
                if ($parseErrors -and $parseErrors.Count -gt 0) {
                    $details = ($parseErrors | Select-Object -First 3 | ForEach-Object {
                            'Zeile {0}: {1}' -f $_.Extent.StartLineNumber, $_.Message
                        }) -join ' | '
                    throw "Syntaxfehler ($($parseErrors.Count)): $details"
                }
            }

            $name = [System.IO.Path]::GetFileNameWithoutExtension($file)   # Scriptname = Dateiname ohne .ps1
            if ($name.Length -gt 200) { throw 'Der Scriptname darf maximal 200 Zeichen lang sein.' }
            $result.ScriptName = $name

            $localHash   = Get-SourceHash -Source $source
            $result.Hash = $localHash

            # Vorhandenes (aktives) Script suchen
            $existing = Get-RepositoryScripts | Where-Object { $_.ScriptName -eq $name } | Select-Object -First 1

            $descToSend = if ($PSBoundParameters.ContainsKey('Description')) {
                $Description
            }
            elseif ($existing) {
                $existing.Description      # Beschreibung beibehalten (die Prozedur würde sonst NULL schreiben)
            }
            else { $null }

            if ($existing) {
                $targetId = $existing.ScriptID
                $action   = 'Update'

                $unchanged = ($existing.HashSHA256 -eq $localHash) -and
                             ([string]$existing.Description -eq [string]$descToSend) -and
                             (-not $Inactive)

                if ($unchanged -and -not $Force) {
                    $result.ScriptID = $existing.ScriptID
                    $result.Version  = $existing.ScriptVersion
                    $result.Action   = 'Unchanged'
                    $result.Status   = 'Skipped'
                    $result.Message  = 'Inhalt und Beschreibung unverändert, kein Upload.'
                    $script:Counters.Unchanged++
                    Write-UploadLog -Message "'$name' (ID $($existing.ScriptID)): unverändert, Version $($existing.ScriptVersion) bleibt bestehen."
                    $result
                    continue
                }
            }
            else {
                $targetId = $null
                $action   = 'Create'
            }

            $target = if ($existing) { "ID $targetId" } else { 'neu' }
            if (-not $PSCmdlet.ShouldProcess("$name ($target)", "Upload ($action)")) {
                $result.Action  = $action
                $result.Status  = 'WhatIf'
                $result.Message = 'Kein Upload (WhatIf).'
                $script:Counters.WhatIf++
                $result
                continue
            }

            Write-UploadLog -Message "Upload gestartet: Datei='$file', Script='$name', Aktion=$action, ScriptID=$target, Hash(lokal)=$localHash, Zeichen=$($source.Length)"

            $res = Invoke-UploadProcedure -ScriptID $targetId -Name $name -Desc $descToSend -Source $source -Active (-not $Inactive)

            $result.ScriptID = $res.ScriptID
            $result.Version  = $res.ScriptVersion
            $result.Action   = $action

            # Integritätsprüfung: serverseitiger Hash muss dem lokalen Hash entsprechen
            if ($res.HashSHA256 -ne $localHash) {
                $result.Status  = 'HashMismatch'
                $result.Message = "Hash-Abweichung: lokal=$localHash, Server=$($res.HashSHA256). Das Script wurde gespeichert, bitte prüfen."
                throw $result.Message
            }

            $result.Status  = 'Success'
            $result.Message = "Version $($res.ScriptVersion) gespeichert."
            if ($action -eq 'Create') { $script:Counters.Created++ } else { $script:Counters.Updated++ }
            Write-UploadLog -Message "Upload erfolgreich: '$name' (ID $($res.ScriptID)), Version $($res.ScriptVersion), Hash bestätigt."
        }
        catch {
            $msg = Get-ErrorText $_
            if ($result.Status -ne 'HashMismatch') { $result.Message = $msg }
            if ($msg -match 'UNIQUE KEY|duplicate key|UQ_PS_Scripts_ScriptName') {
                $result.Message += ' Hinweis: Ein Script mit diesem Namen existiert vermutlich inaktiv. Inaktive Scripte sind für diese Rolle nicht sichtbar, bitte Administrator kontaktieren.'
            }
            $script:HadError = $true
            $script:Counters.Failed++
            Write-UploadLog -Level ERROR -Message "Upload fehlgeschlagen: '$file' -> $($result.Message)"
        }

        $result
    }
}

end {
    if ($script:Conn) {
        $script:Conn.Dispose()
    }

    $c = $script:Counters
    Write-UploadLog -Message ("Zusammenfassung: neu={0}, aktualisiert={1}, unverändert={2}, WhatIf={3}, fehlgeschlagen={4}" -f `
            $c.Created, $c.Updated, $c.Unchanged, $c.WhatIf, $c.Failed)

    if ($script:HadError) { exit 1 }
}