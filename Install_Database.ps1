<#
.SYNOPSIS
    Creates two domain groups, adds the current user to them and runs a SQL file
    against a MS SQL Server. Replaces the placeholders #DBNAME# and #DOMAIN#.

.DESCRIPTION
    Shows basic information first and lets the user abort.
    Then prompts for:
      - SQL file (must exist, extension .SQL)
      - SQL Server (instance)
      - Authentication (Windows or SQL Server)
      - User name / password (SQL Server authentication only)
      - New database name (default: PSRepository)
      - Domain short name (NetBIOS name, detected automatically)
      - Optional: OU in which the domain groups are created

    Then, after a summary and a confirmation:
      1. Connects to the SQL Server and checks that the database does not exist
      2. Creates the domain groups GG_PS_Executors and GG_PS_Uploaders
         (global security groups; existing groups are reused)
      3. Adds the current user to both groups
      4. Runs the SQL file
      5. Replaces #DBNAME# (database name) and #DBSERVER# (SQL Server) in
         PS_Download.ps1, PS_Execute.ps1 and PS_Upload.ps1
         (same folder as this script; the files are changed in place)

    Placeholders in the SQL file:
      #DBNAME#  -> name of the new database
      #DOMAIN#  -> domain short name, e.g. #DOMAIN#\GG_PS_Uploaders -> CONTOSO\GG_PS_Uploaders

    GO batch separators in the SQL script are supported.
    No additional modules are required (uses System.Data.SqlClient and ADSI).

.PARAMETER SqlFile
    Path to the SQL file (must exist and have the extension .SQL). Relative paths
    are resolved relative to the folder of this script. If omitted, the default
    is offered at a prompt, which repeats until an existing .SQL file is entered.

.PARAMETER TrustServerCertificate
    Trusts the server certificate (e.g. for self-signed certificates).

.EXAMPLE
    .\Invoke-SqlFile.ps1 -SqlFile .\SQL\Database\Create_Database.SQL
#>

[CmdletBinding()]
param(
    [string]$SqlFile = '\SQL\Database\Create_Database.SQL',
    [switch]$TrustServerCertificate
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.DirectoryServices

$PlaceholderDb     = '#DBNAME#'
$PlaceholderDomain = '#DOMAIN#'
$PlaceholderServer = '#DBSERVER#'
$GroupNames        = @('GG_PS_Executors', 'GG_PS_Uploaders')

# ---------------------------------------------------------------
# Helper: Read-Host with default value
# ---------------------------------------------------------------
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

# ---------------------------------------------------------------
# Helper: Yes/No question, returns $true for Yes
# ---------------------------------------------------------------
function Read-YesNo {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [string]$Default = 'Y'
    )
    $answer = ''
    while ($answer -notin @('Y', 'N')) {
        $answer = (Read-HostWithDefault -Prompt "$Prompt (Y)es / (N)o" -Default $Default).Substring(0, 1).ToUpper()
    }
    return ($answer -eq 'Y')
}

# ---------------------------------------------------------------
# Helper: Build an LDAP path (bound to one domain controller)
# ---------------------------------------------------------------
function Get-LdapPath {
    param(
        [Parameter(Mandatory)]$Domain,
        [Parameter(Mandatory)][string]$DistinguishedName
    )
    # A "/" inside a DN must be escaped in ADSI paths
    return $Domain.LdapPrefix + $DistinguishedName.Replace('/', '\/')
}

# ---------------------------------------------------------------
# Helper: Get domain information (naming context, NetBIOS name, DC)
# ---------------------------------------------------------------
function Get-DomainInfo {
    try {
        $rootDse   = [ADSI]'LDAP://RootDSE'
        $defaultNC = [string]$rootDse.Properties['defaultNamingContext'][0]
        $configNC  = [string]$rootDse.Properties['configurationNamingContext'][0]
        $dcName    = [string]$rootDse.Properties['dnsHostName'][0]
    }
    catch {
        throw "Active Directory is not reachable. This computer must be joined to a domain (or connected to it). Details: $($_.Exception.Message)"
    }
    if ([string]::IsNullOrWhiteSpace($defaultNC)) {
        throw 'Active Directory is not reachable: the default naming context could not be determined.'
    }

    # Use the same domain controller for all following operations (avoids replication delays)
    $prefix = if ($dcName) { "LDAP://$dcName/" } else { 'LDAP://' }

    # NetBIOS (short) name of the domain, taken from the Partitions container
    $netBios = $null
    try {
        $partitions = [ADSI]($prefix + "CN=Partitions,$configNC")
        $searcher = New-Object System.DirectoryServices.DirectorySearcher($partitions)
        $searcher.Filter      = "(&(objectClass=crossRef)(nCName=$defaultNC))"
        $searcher.SearchScope = 'OneLevel'
        [void]$searcher.PropertiesToLoad.Add('nETBIOSName')
        $result = $searcher.FindOne()
        if ($result -and $result.Properties['netbiosname'].Count -gt 0) {
            $netBios = [string]$result.Properties['netbiosname'][0]
        }
    }
    catch { }
    if (-not $netBios) { $netBios = $env:USERDOMAIN }

    return [pscustomobject]@{
        DefaultNamingContext = $defaultNC
        NetBiosName          = $netBios
        LdapPrefix           = $prefix
    }
}

# ---------------------------------------------------------------
# Helper: Find a domain group by its sAMAccountName ($null if missing)
# ---------------------------------------------------------------
function Find-DomainGroup {
    param(
        [Parameter(Mandatory)]$Domain,
        [Parameter(Mandatory)][string]$Name
    )
    $root = [ADSI](Get-LdapPath -Domain $Domain -DistinguishedName $Domain.DefaultNamingContext)
    $searcher = New-Object System.DirectoryServices.DirectorySearcher($root)
    $searcher.Filter = "(&(objectCategory=group)(sAMAccountName=$Name))"
    $result = $searcher.FindOne()
    if ($result) { return $result.GetDirectoryEntry() }
    return $null
}

# ---------------------------------------------------------------
# Helper: Create a global security group
# ---------------------------------------------------------------
function New-DomainGroup {
    param(
        [Parameter(Mandatory)]$Domain,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$ContainerDn,
        [string]$Description
    )
    $container = [ADSI](Get-LdapPath -Domain $Domain -DistinguishedName $ContainerDn)
    $group = $container.psbase.Children.Add("CN=$Name", 'group')
    $group.psbase.Properties['sAMAccountName'].Value = $Name
    $group.psbase.Properties['groupType'].Value      = [int]-2147483646   # global + security-enabled
    if ($Description) { $group.psbase.Properties['description'].Value = $Description }
    $group.psbase.CommitChanges()
    return $group
}

# ---------------------------------------------------------------
# Helper: Add a user (by DN) to a group, returns $true if added
# ---------------------------------------------------------------
function Test-GroupMember {
    param(
        [Parameter(Mandatory)]$Group,
        [Parameter(Mandatory)][string]$UserDn
    )
    # Bind a fresh object and reload 'member' from the server.
    # (The property cache of an already bound object is not updated automatically.)
    $fresh = [ADSI]([string]$Group.psbase.Path)
    $fresh.psbase.RefreshCache(@('member'))
    return ($fresh.psbase.Properties['member'] -contains $UserDn)
}

function Add-UserToGroup {
    param(
        [Parameter(Mandatory)]$Group,
        [Parameter(Mandatory)]$Domain,
        [Parameter(Mandatory)][string]$UserDn
    )
    # Note: ADSI objects hide the .NET members (CommitChanges, Properties, ...) behind "psbase"
    if (Test-GroupMember -Group $Group -UserDn $UserDn) { return $false }

    $groupPath = [string]$Group.psbase.Path
    $firstError  = $null
    $secondError = $null

    # Attempt 1: add the user DN to the 'member' attribute
    try {
        $entry = [ADSI]$groupPath
        [void]$entry.psbase.Properties['member'].Add($UserDn)
        $entry.psbase.CommitChanges()
    }
    catch { $firstError = $_.Exception.Message }

    # Attempt 2 (fallback): native IADsGroup.Add with the ADsPath of the user
    if (-not (Test-GroupMember -Group $Group -UserDn $UserDn)) {
        try {
            $userPath = Get-LdapPath -Domain $Domain -DistinguishedName $UserDn
            $entry = [ADSI]$groupPath
            [void]$entry.psbase.Invoke('Add', $userPath)
        }
        catch { $secondError = $_.Exception.Message }
    }

    # Verify against the server - never report success without proof
    if (-not (Test-GroupMember -Group $Group -UserDn $UserDn)) {
        throw "The user '$UserDn' is not a member of the group after adding. Errors: [1] $firstError [2] $secondError"
    }
    return $true
}

# ---------------------------------------------------------------
# Helper: Read a text file and remember its encoding (BOM aware),
# so that it can be written back unchanged apart from the replacement
# ---------------------------------------------------------------
function Get-TextFile {
    param([Parameter(Mandatory)][string]$Path)

    $bytes  = [System.IO.File]::ReadAllBytes($Path)
    $enc    = $null
    $offset = 0

    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $enc = New-Object System.Text.UTF8Encoding($true)            # UTF-8 with BOM
        $offset = 3
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        $enc = New-Object System.Text.UnicodeEncoding($false, $true) # UTF-16 LE with BOM
        $offset = 2
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        $enc = New-Object System.Text.UnicodeEncoding($true, $true)  # UTF-16 BE with BOM
        $offset = 2
    }
    else {
        # No BOM: UTF-8 if the bytes are valid UTF-8, otherwise the system ANSI code page
        $strict = New-Object System.Text.UTF8Encoding($false, $true)
        try {
            [void]$strict.GetString($bytes)
            $enc = New-Object System.Text.UTF8Encoding($false)
        }
        catch {
            $enc = [System.Text.Encoding]::Default
        }
    }

    return [pscustomobject]@{
        Text     = $enc.GetString($bytes, $offset, $bytes.Length - $offset)
        Encoding = $enc
    }
}
# ---------------------------------------------------------------
# Helper: Reset Installation ( Undo changes on Script Files)
# 
# ---------------------------------------------------------------
function Set-ToOriginal{
  
    $Source = Join-Path $scriptPath -ChildPath "Scripts\*.*"
    Copy-Item  -Path $Source -Destination $scriptPath -Force
}
try {

    if ($PSScriptRoot) {
        $scriptPath = $PSScriptRoot
    }
    elseif ($MyInvocation.MyCommand.Path) {
        $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    elseif ($isISE -and $psISE.CurrentFile -and -not $psISE.CurrentFile.IsUntitled) {
        # F8 in der ISE: Pfad der Datei im aktiven Tab
        $scriptPath = Split-Path -Parent $psISE.CurrentFile.FullPath
    }
    else {
        $scriptPath = (Get-Location).Path
    }

    # -----------------------------------------------------------
    # Reset chanches on Files
    # -----------------------------------------------------------

    Set-ToOriginal
    # -----------------------------------------------------------
    # 0. Show Basic Infos
    # -----------------------------------------------------------
    Write-host "You need: "
    Write-host "  - A running MS SQL Server (Express is possible)" -ForegroundColor Yellow
    Write-host "  - This Account has the Rights to create Domain Groups" -ForegroundColor Yellow
    Write-host "  - The database does not exist yet" -ForegroundColor Yellow
    Write-Host ''

    # Give the user the chance to abort (Ctrl+C also works at any prompt)
    if (-not (Read-YesNo -Prompt 'Do you want to continue?')) {
        Write-Host 'Aborted by user. Nothing was changed.' -ForegroundColor Yellow
        exit 0
    }
    Write-Host ''

    # -----------------------------------------------------------
    # 1. Determine the SQL file
    # -----------------------------------------------------------
    # Determine the folder of this script (base for relative paths)
    $scriptDir = $null
    if ($PSScriptRoot) {
        $scriptDir = $PSScriptRoot
    }
    elseif ($MyInvocation.MyCommand.Path) {
        $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    else {
        # Happens when the script was not started as a file (e.g. ISE selection with F8 or copy & paste)
        $scriptDir = (Get-Location).Path
        Write-Warning "Location of the script could not be determined - using current directory: $scriptDir"
    }

    # If -SqlFile was passed, it is validated once (error on failure).
    # Otherwise the default is offered at a prompt until an existing .SQL file is entered.
    $askForFile     = -not $PSBoundParameters.ContainsKey('SqlFile')
    $sqlFileDefault = $SqlFile
    $validSqlFile   = $scriptPath + $sqlFileDefault

    while (-not $validSqlFile) {
        if ($askForFile) {
            $SqlFile = Read-HostWithDefault -Prompt 'Path to the SQL file (relative to the script folder)' -Default $sqlFileDefault
        }

        # Resolve relative paths (also with a leading \ or /) relative to the script folder
        $candidate = $SqlFile
        if ($candidate -match '^[\\/](?![\\/])') {
            # "\SQL\..." -> "SQL\..." (meant as relative, not as drive root)
            $candidate = $candidate.TrimStart('\', '/')
        }
        if (-not [System.IO.Path]::IsPathRooted($candidate)) {
            $candidate = Join-Path -Path $scriptDir -ChildPath $candidate
        }

        # Check: the file must exist and must have the extension .SQL (case-insensitive)
        $problem = $null
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            $problem = "File not found: $candidate"
        }
        elseif ([System.IO.Path]::GetExtension($candidate) -ine '.sql') {
            $problem = "Not a .SQL file: $candidate"
        }

        if ($problem) {
            if (-not $askForFile) { throw $problem }
            Write-Host $problem -ForegroundColor Red
            Write-Host 'Please enter the path of an existing .SQL file (Ctrl+C to abort).' -ForegroundColor Yellow
            continue
        }

        $validSqlFile = $candidate
    }

    $SqlFile = $validSqlFile
    Write-Host "SQL file: $SqlFile" -ForegroundColor DarkGray

    # -----------------------------------------------------------
    # 2. Ask for connection information
    # -----------------------------------------------------------
    $server = Read-HostWithDefault -Prompt 'SQL Server (e.g. localhost or SERVER\INSTANCE)' -Default 'localhost'

    $authMode = ''
    while ($authMode -notin @('W', 'S')) {
        $authMode = (Read-HostWithDefault -Prompt 'Authentication: (W)indows or (S)QL Server' -Default 'W').ToUpper()
    }

    $user = $null
    $securePassword = $null
    if ($authMode -eq 'S') {
        $user = Read-HostWithDefault -Prompt 'User name'
        $securePassword = Read-Host -Prompt 'Password' -AsSecureString
    }

    # -----------------------------------------------------------
    # 3. Ask for the new database name (default: PSRepository)
    # -----------------------------------------------------------
    $dbName = Read-HostWithDefault -Prompt 'Name of the new database' -Default 'PSRepository'

    # The name is inserted directly into the SQL -> allow safe characters only
    if ($dbName -notmatch '^[A-Za-z0-9_\-\.]{1,128}$') {
        throw "Invalid database name '$dbName'. Allowed: letters, digits, '_', '-' and '.' (max. 128 characters)."
    }

    # -----------------------------------------------------------
    # 4. Domain information, OU for the groups and current user
    # -----------------------------------------------------------
    Write-Host ''
    Write-Host 'Reading domain information ...' -ForegroundColor Cyan
    $domain = Get-DomainInfo

    $domainShort = Read-HostWithDefault -Prompt 'Domain short name (NetBIOS name)' -Default $domain.NetBiosName

    # The name is inserted directly into the SQL -> allow safe characters only
    if ($domainShort -notmatch '^[A-Za-z0-9_\-\.]{1,15}$') {
        throw "Invalid domain short name '$domainShort'. Allowed: letters, digits, '_', '-' and '.' (max. 15 characters)."
    }

    # Optional OU (distinguished name). Empty = default container CN=Users.
    # The OU is only used for groups that do not exist yet.
    $containerDn = $null
    while (-not $containerDn) {
        $ouInput = Read-Host "OU for the new groups as distinguished name (e.g. OU=Groups,$($domain.DefaultNamingContext)) - leave empty for the default container"
        $ouInput = $ouInput.Trim()
        $candidateDn = if ([string]::IsNullOrWhiteSpace($ouInput)) { "CN=Users,$($domain.DefaultNamingContext)" } else { $ouInput }

        $ouExists = $false
        try { $ouExists = [ADSI]::Exists((Get-LdapPath -Domain $domain -DistinguishedName $candidateDn)) } catch { }

        if ($ouExists) {
            $containerDn = $candidateDn
        }
        elseif ([string]::IsNullOrWhiteSpace($ouInput)) {
            throw "The default container '$candidateDn' does not exist. Please run the script again and enter an OU."
        }
        else {
            Write-Host "OU/container not found: $candidateDn" -ForegroundColor Red
            Write-Host 'Please enter an existing distinguished name or leave empty (Ctrl+C to abort).' -ForegroundColor Yellow
        }
    }

    # Current user (must be a domain account to be added to the groups)
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $userSid  = $identity.User.Value
    try {
        $userEntry = [ADSI](Get-LdapPath -Domain $domain -DistinguishedName "<SID=$userSid>")
        $userDn    = [string]$userEntry.Properties['distinguishedName'][0]
    }
    catch {
        throw "The current user '$($identity.Name)' could not be found in Active Directory (local account?). Details: $($_.Exception.Message)"
    }
    if ([string]::IsNullOrWhiteSpace($userDn)) {
        throw "The current user '$($identity.Name)' could not be found in Active Directory."
    }

    # -----------------------------------------------------------
    # 5. Read the SQL & PS1 File(s) and replace the placeholders
    # -----------------------------------------------------------
    $sql = Get-Content -LiteralPath $SqlFile -Raw -Encoding UTF8
    if ($sql -notmatch [regex]::Escape($PlaceholderDb)) {
        Write-Warning "The placeholder $PlaceholderDb was not found in '$SqlFile'."
    }
    if ($sql -notmatch [regex]::Escape($PlaceholderDomain)) {
        Write-Warning "The placeholder $PlaceholderDomain was not found in '$SqlFile'."
    }
    $sql = $sql.Replace($PlaceholderDb, $dbName).Replace($PlaceholderDomain, $domainShort)

    # Split into batches at GO lines
    $batches = [regex]::Split($sql, '(?im)^[ \t]*GO[ \t]*\r?$') |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    # PowerShell scripts in the same folder that also contain the placeholder.
    # They are checked and prepared here, but only written after the SQL ran successfully.
    $psScriptNames = @('PS_Download.ps1', 'PS_Execute.ps1', 'PS_Upload.ps1')
    # The placeholder sits inside single-quoted strings in the PS scripts -> double any single quote
    $serverForPs   = $server.Replace("'", "''")
    $psScripts = @(foreach ($name in $psScriptNames) {
        $path = Join-Path -Path $scriptDir -ChildPath $name
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "PowerShell script not found: $path"
        }
        $file   = Get-TextFile -Path $path
        $hasDb  = $file.Text.Contains($PlaceholderDb)
        $hasSrv = $file.Text.Contains($PlaceholderServer)
        if (-not $hasDb)  { Write-Warning "The placeholder $PlaceholderDb was not found in '$path'." }       
        [pscustomobject]@{
            Path     = $path
            NewText  = $file.Text.Replace($PlaceholderDb, $dbName).Replace($PlaceholderServer, $serverForPs)
            Encoding = $file.Encoding
            Changed  = ($hasDb -or $hasSrv)
        }
    })

    # -----------------------------------------------------------
    # 6. Summary and confirmation
    # -----------------------------------------------------------
    $authText  = if ($authMode -eq 'W') { 'Windows authentication' } else { 'SQL authentication, user ' + $user }
    $groupList = ($GroupNames | ForEach-Object { "$domainShort\$_" }) -join ', '

    Write-Host ''
    Write-Host 'Summary' -ForegroundColor Cyan
    Write-Host "  SQL file          : $SqlFile"
    Write-Host "  SQL Server        : $server ($authText)"
    Write-Host "  New database      : $dbName"
    Write-Host "  Domain            : $domainShort"
    Write-Host "  Domain groups     : $groupList"
    Write-Host "  Groups created in : $containerDn (only if they do not exist yet)"
    Write-Host "  Current user      : $($identity.Name) (will be added to both groups)"
    Write-Host "  PowerShell scripts: $($psScriptNames -join ', ') ($PlaceholderDb and $PlaceholderServer will be replaced)"
    Write-Host ''
    if (-not (Read-YesNo -Prompt 'Start now?')) {
        Write-Host 'Aborted by user. Nothing was changed.' -ForegroundColor Yellow
        exit 0
    }

    # -----------------------------------------------------------
    # 7. Open the SQL connection (to the master database first)
    # -----------------------------------------------------------
    $csb = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $csb['Data Source']      = $server
    $csb['Initial Catalog']  = 'master'
    $csb['Application Name'] = 'Invoke-SqlFile'
    if ($TrustServerCertificate) { $csb['TrustServerCertificate'] = $true }

    if ($authMode -eq 'W') {
        $csb['Integrated Security'] = $true
        $connection = New-Object System.Data.SqlClient.SqlConnection($csb.ConnectionString)
    }
    else {
        $securePassword.MakeReadOnly()
        $credential = New-Object System.Data.SqlClient.SqlCredential($user, $securePassword)
        $connection = New-Object System.Data.SqlClient.SqlConnection($csb.ConnectionString, $credential)
    }

    # Show PRINT output and messages from the server
    $connection.add_InfoMessage({
        param($sender, $e)
        Write-Host "  SQL: $($e.Message)" -ForegroundColor DarkGray
    })
    $connection.FireInfoMessageEventOnUserErrors = $false

    Write-Host ''
    Write-Host "Connecting to '$server' ..." -ForegroundColor Cyan
    $connection.Open()
    Write-Host "Connected (SQL Server $($connection.ServerVersion))." -ForegroundColor Green

    # The database must not exist yet (checked before anything is changed in AD)
    $checkCmd = $connection.CreateCommand()
    $checkCmd.CommandText = 'SELECT DB_ID(@name)'
    [void]$checkCmd.Parameters.AddWithValue('@name', $dbName)
    $dbId = $checkCmd.ExecuteScalar()
    $checkCmd.Dispose()
    if ($null -ne $dbId -and $dbId -isnot [System.DBNull]) {
        throw "The database '$dbName' already exists. Nothing was changed. Choose another name or remove the database first."
    }

    # -----------------------------------------------------------
    # 8. Create the domain groups and add the current user
    # -----------------------------------------------------------
    Write-Host ''
    Write-Host 'Creating domain groups ...' -ForegroundColor Cyan
    foreach ($groupName in $GroupNames) {
        try {
            $group = Find-DomainGroup -Domain $domain -Name $groupName
            if ($group) {
                Write-Host "  Group '$groupName' already exists - using the existing group."
            }
            else {
                $group = New-DomainGroup -Domain $domain -Name $groupName -ContainerDn $containerDn -Description "Created by Invoke-SqlFile.ps1 for database $dbName"
                Write-Host "  Group '$groupName' created in '$containerDn'." -ForegroundColor Green
            }

            if (Add-UserToGroup -Group $group -Domain $domain -UserDn $userDn) {
                Write-Host "  User '$($identity.Name)' added to '$groupName'." -ForegroundColor Green
            }
            else {
                Write-Host "  User '$($identity.Name)' is already a member of '$groupName'."
            }
        }
        catch {
            throw "Failed to create or change the domain group '$groupName': $($_.Exception.Message) - does this account have the right to create groups in '$containerDn'?"
        }
    }

    # -----------------------------------------------------------
    # 9. Execute the batches
    # -----------------------------------------------------------
    Write-Host ''
    Write-Host "Running '$SqlFile' (database: $dbName, $($batches.Count) batch(es)) ..." -ForegroundColor Cyan

    $i = 0
    foreach ($batch in $batches) {
        $i++
        $cmd = $connection.CreateCommand()
        $cmd.CommandText    = $batch
        $cmd.CommandTimeout = 0   # no timeout
        try {
            [void]$cmd.ExecuteNonQuery()
        }
        catch {
            throw "Error in batch $i of $($batches.Count): $($_.Exception.Message)"
        }
        finally {
            $cmd.Dispose()
        }
    }

    # -----------------------------------------------------------
    # 10. Replace the placeholder in the PowerShell scripts
    # -----------------------------------------------------------
    Write-Host ''
    Write-Host 'Updating PowerShell scripts ...' -ForegroundColor Cyan
    foreach ($ps in $psScripts) {
        if (-not $ps.Changed) {
            Write-Host "  Skipped (no placeholder found): $($ps.Path)" -ForegroundColor Yellow
            continue
        }
        try {
            [System.IO.File]::WriteAllText($ps.Path, $ps.NewText, $ps.Encoding)
        }
        catch {
            throw "The SQL script was executed, but '$($ps.Path)' could not be updated: $($_.Exception.Message)"
        }
        Write-Host "  Updated: $($ps.Path)" -ForegroundColor Green
    }


    Write-Host ''
    Write-Host "Upload Upload Script and Download Script" -ForegroundColor Green

& "$scriptPath\PS_Upload.ps1" `
    -FilePath "$scriptPath\PS_Upload.ps1" `
    -Description "Script to Upload other Script (useable with PS_Execute)"

& "$scriptPath\PS_Upload.ps1" `
    -FilePath "$scriptPath\PS_Download.ps1" `
    -Description "Script to Download Scripts from Database (useable with PS_Execute)"


    Write-Host ''
    Write-Host "Done. The script was executed successfully (database: $dbName)." -ForegroundColor Green
    Write-Host 'Note: Log off and on again so that the new group membership of the current user takes effect.' -ForegroundColor Yellow

 
 
 
    
    Start-Process "$scriptPath\Shortintroduktion.pdf"
}
catch {
    Write-Host ''
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
finally {
    if ($connection) {
        if ($connection.State -ne 'Closed') { $connection.Close() }
        $connection.Dispose()
    }
}