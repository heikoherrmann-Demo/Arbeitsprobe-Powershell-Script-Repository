<#
.SYNOPSIS
    PRTG Sensor for PS Repository.
    Replaces the placeholders #DBNAME# and #DATABASENAME#.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '')]
param()

$dbServer = "#DBNAME#"
$dbName = "#DATABASENAME#"

#region Basic Database Functions
function db-ConverttoDbTypeDatatypes {
    # This function resturns Some Conversion functions:
    # as Array:
    # ret[0] = Type of Object
    # ret[1] = Type of Database Object
    # ret[2] = Value for Insert Statement
    # ret[3] = Value for Update Statement

    param(
    [Parameter(Mandatory=$true)][Object]$Object,
    [Parameter(Mandatory=$true)][String]$Member
    )
    $ret = @()
    # [0] = Object Type
    if ( $null -eq $Object.($Member) )
    {
        # Etwas Komplex: Aus Member die Definition enthält die VarType
        $varType = (((($Object | Get-Member | where-object {$_.name -eq $Member}).Definition).split(' '))[0]).Split('.')[1]
    }
    else
    {
        $varType  = ($Object.($Member)).GetType().Name
    }
    $ret += $varType
    switch ($varType) 
    { 
        "Boolean" { $ret += "[bit]"
                    if ($Object.($Member)){$ret += "0"}            else {$ret += "1"}
                    if ($Object.($Member)){$ret += "$($Member)=0"} else {$ret += "$($Member)=1"}
                    }
                    
        "Integer" { $ret += "[bigint]"
                    $ret += "$($Object.($Member))" 
                    $ret += "$($Member)=$($Object.($Member))"} 
        "String"  { $ret += "[varchar](200)"
                    $ret += "'$($Object.($Member))'"
                    $ret += "$($Member)='$($Object.($Member))'"    } 
        default   { $ret += "[varchar](200)"
                    $ret += "'$($Object.($Member))'"
                    $ret += "$($Member)='$($Object.($Member))'"    }
    }
    return $ret
}
function db-open {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SQLServer,

        [Parameter(Mandatory=$true)]
        [string]$Databasename,

        [Parameter(Mandatory=$false)]
        [string]$Username
    )

    # Create SQL Connection
    # ---------------------
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection

    # Windows Authentication
    $SqlConnection.ConnectionString = "Server=$SQLServer;Database=$Databasename;Integrated Security=True;"

    # Open Connection
    # ----------------
    try {
        $SqlConnection.Open()
        return $SqlConnection
    }
    catch {
        $SqlConnection.Dispose()
        throw "Database connection failed: $($_.Exception.Message)"
    }
}

function db-close {
    param(
        [Parameter(Mandatory=$true)][System.Data.SqlClient.SqlConnection]$Connection
        )

    $Connection.Close()
}
function db-Execute {
    param(
    [Parameter(Mandatory=$true)][string]$SqlCommand,
    [Parameter(Mandatory=$true)][System.Data.SqlClient.SqlConnection]$Connection
    )
    $cmd = New-Object System.Data.SqlClient.SqlCommand
    $cmd.connection = $Connection
    $cmd.commandtext = $SqlCommand
    return $cmd.ExecuteNonQuery()
}
function db-get {
    param(
        [Parameter(Mandatory=$true)][string]$SqlQuery,
        [Parameter(Mandatory=$true)][System.Data.SqlClient.SqlConnection]$Connection
    )
    #  Query Database
    # ---------------
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $SqlCmd.CommandText = $SqlQuery
    $SqlCmd.Connection = $Connection
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlAdapter.SelectCommand = $SqlCmd
    $DataSet = New-Object System.Data.DataSet
    $DataSet.Clear()
    $SqlAdapter.Fill($DataSet)
    #$DataSet.Tables[0] 
    return [System.Data.DataSet] $DataSet
    }
function db-insert {
    param(
    [Parameter(Mandatory=$true)][System.Data.SqlClient.SqlConnection]$Connection,
    [Parameter(Mandatory=$true)][Object[]]$Objectlist,
    [Parameter(Mandatory=$true)][String]$PrimaryKey,
    [Parameter(Mandatory=$true)][String]$Tablename
    )
    $PropertyList = ("Property","NoteProperty")

    # ##############################Functions############################################
    function db-createInsert {
        param(
        [Parameter(Mandatory=$true)][System.Object]$Object,
        [Parameter(Mandatory=$true)][string]$Tablename
        )
        # $Object |Get-Member
        $fieldsnum = 0
        $sqlfield = "insert into $Tablename ( "
        $sqlvalues = "VALUES ( "
        foreach( $p in ($Object |Get-Member))
        {
            if ($PropertyList -contains $p.MemberType)
            {
                $sqlfield += "$($p.name) ,"
                $sqlvalues += (db-ConverttoDbTypeDatatypes -Object $Object -Member $p.name)[2]+" ,"
                $fieldsnum  += 1
            }
        }
        if ( $fieldsnum -gt 0 )
        {
           $sqlfield += "LastChange ) "
           $sqlvalues += "'{0}'"+ ") "
        } 
        
        
        return $($sqlfield  + $sqlvalues )
    }
    function db-createUpdate {
        param(
        [Parameter(Mandatory=$true)][System.Object]$Object,
        [Parameter(Mandatory=$true)][string]$PrimaryKeyField,
        [Parameter(Mandatory=$true)][string]$PrimaryKeyValue,
        [Parameter(Mandatory=$true)][string]$Tablename
        )
        $sql = "Update $Tablename Set `r`n   "
        foreach( $p in ($Object |Get-Member))
        {
            # TODO: Prüfen ob feld auch in DB existiert
            if ($PropertyList -contains $p.MemberType)
            {
                $sql += "[$($p.name)] = "
                $sql +=  (db-ConverttoDbTypeDatatypes -Object $Object -Member $p.name)[2]+",`r`n   "
            }
        }
        $sql += "[LastChange] = '{0}'`r`n"
        $sql += "where $PrimaryKeyField="+$PrimaryKeyValue
        
        return $sql
    }
    # ###################################################################################
    # Common Checks
    # 1) Check if Object[] is empty
    # 2) Primary Key in Object Properties?


    # Check if Table allready exists
    #    -> when not: Create Table
    $sql = "SELECT * FROM information_schema.tables WHERE TABLE_TYPE='BASE TABLE' and TABLE_Name = '"+$Tablename+"'"
    $table = db-get -SqlQuery $sql -Connection $Connection
    if ( $table[0] -eq 0 )
    {
        $o = $Objectlist[0]
        # Create a new Table
        ####################
        $sql = "CREATE TABLE [$Tablename] (`r`n "
        foreach( $p in ($o |Get-Member))
        {
            if ($PropertyList -contains $p.MemberType)
            {
                $sql += "  [$($p.name)] "
                $sql += $((db-ConverttoDbTypeDatatypes -Object $o -Member $p.name)[1])
                if ($p.name -eq $PrimaryKey)
                {
                    # Null not allowed at Primary Key
                    $sql += " NOT NULL,`r`n"
                }
                else
                {
                    $sql += " NULL,`r`n"
                }
            }
        }
        $sql += "[LastChange] [datetime] NOT NULL `r`n"
        $sql += "CONSTRAINT ["+$Tablename+"_1_1] PRIMARY KEY CLUSTERED "
        $sql += "( [$PrimaryKey] ASC ) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
        $sql += ") ON [PRIMARY]"
        $res = db-Execute -Connection $Connection -SqlCommand $sql                 
    }

    if ( (db-getPrimaryKey -Tablename $Tablename -Connection $Connection) -ne $PrimaryKey)
    {
        write-host "Primary keys didn't Match:" -ForegroundColor red
        write-host " PK in Function: $PrimaryKey" -ForegroundColor red
        write-host " PK in Table   : $(db-getPrimaryKey -Tablename $Tablename -Connection $Connection)" -ForegroundColor red
        return
    }
    # insert Date is for alle Entries the Same
    $CreateDate = Get-Date -format "yyyy-MM-dd HH:mm:ss"
    #Write-Host "$Tablename : $CreateDate ($($Objectlist.count))"
    # For each Object in Object Liste check if it allready exists
    foreach( $o in $Objectlist)
    {
        $vals = db-ConverttoDbTypeDatatypes -Object $o -Member $PrimaryKey       
        $sql = "select * from $Tablename where $PrimaryKey= $($vals[2])"
        #write-host "primary key: $sql"
        $tab = db-get -Connection $Connection -SqlQuery $sql
        if ( $tab[0] -gt 0 )
        {
            #write-host " Update $PrimaryKey = $pk" -ForegroundColor Yellow
            $sql = (db-createUpdate -Object $o -Tablename $tablename -PrimaryKeyField $PrimaryKey -PrimaryKeyValue $vals[2]) -f $CreateDate
        }
        else
        {
            #write-host " Insert $PrimaryKey = $pk" -ForegroundColor green
            $sql = (db-createInsert -Object $o -Tablename $tablename ) -f $CreateDate
        }
        # Execute SQL
        $res = db-Execute -Connection $Connection -SqlCommand $sql

    }

    # last Step delete all Set the are not from $CreateDate
    $sql = "delete from $Tablename where LastChange <> '$CreateDate'"
    $res = db-Execute -Connection $Connection -SqlCommand $sql
    return $res
}
function db-getPrimaryKey{
    param(
    [Parameter(Mandatory=$true)][string]$Tablename,
    [Parameter(Mandatory=$true)][System.Data.SqlClient.SqlConnection]$Connection
    )

    $sql = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc JOIN INFORMATION_SCHEMA.CONSTRAINT_COLUMN_USAGE ccu ON tc.CONSTRAINT_NAME = ccu.Constraint_name  WHERE tc.CONSTRAINT_TYPE = 'Primary Key' and tc.Table_Name = '"+$Tablename+"'"

    $table = db-get -SqlQuery $sql -Connection $Connection
    return $table[1].Tables[0].Rows[0].COLUMN_NAME
}
function db-TestTableExists {
    param(
    [Parameter(Mandatory=$true)][string]$Tablename,
    [Parameter(Mandatory=$true)][System.Data.SqlClient.SqlConnection]$Connection
    )

$sql = @"
    IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES 
                WHERE TABLE_NAME = N'$Tablename')
        Select  -1
    ELSE
        Select 0
"@

    $cmd = New-Object System.Data.SqlClient.SqlCommand
    $cmd.connection = $Connection
    $cmd.commandtext = $sql
    return [bool] $cmd.ExecuteScalar()
}
function db-Recordcount {
    param(
    [Parameter(Mandatory=$true)][string]$SQL,
    [Parameter(Mandatory=$true)][System.Data.SqlClient.SqlConnection]$Connection
    )
    $query = "select count(*) from ($sql) as Counter"
    $cmd = New-Object System.Data.SqlClient.SqlCommand
    $cmd.connection = $Connection
    $cmd.commandtext = $query
    return $cmd.ExecuteScalar()
}
#endregion


$con = db-open -SQLServer $dbServer -Databasename $dbName
$sql = "PS_GetScripts"

WRITE-HOST "<PRTG>"

foreach( $record in (db-get -SqlQuery $sql -con $con).tables.Rows)
{
    # Get Last Result from Script
    $sql = "PS_GetLastExecution " + $record.ScriptID
    $result = db-get -SqlQuery $sql -con $con
    if ($result.Tables[0].Rows.Count -eq 0) 
    {
        $resultmessage = "n/a"
    }
    else
    {
        $resultmessage = $result.tables.Rows[0].ExecutionStatus
    }

    # Show in PRTG
    Write-Host "  <result>"
    Write-Host "    <channel>$($record.ScriptName)</channel>"
    if ($resultmessage -in @("success", "n/a"))
    {
        Write-Host "    <value>0</value>"
    }
    else
    {
        Write-Host "    <value>1</value>"
    }    
    Write-Host "    <valueLookup>prtg.customlookups.scriptstatus</valueLookup>"
    Write-Host "  </result>"

}

WRITE-HOST "</PRTG>"
db-close -Connection $con