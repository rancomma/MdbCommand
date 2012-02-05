#requires -version 2.0

#MdbCommand
#Author rancomma@gmail.com twitter @rancomma

#Used Aliases
#  select = Select-Object
#    sort = Sort-Object
#       % = ForEach-Object
#     iex = Invoke-Expression

#Script Variable
[System.Data.OleDb.OleDbConnection]${script:MdbCommand.Connection} = $null
[System.Data.OleDb.OleDbTransaction]${script:MdbCommand.Transaction} = $null
[System.Data.OleDb.OleDbCommand]${script:MdbCommand.Command} = $null
[string]${script:MdbCommand.FilePath} = $null
[int]${script:MdbCommand.RecordsAffected} = $null
[int]${script:MdbCommand.BatchTotalRecordsAffected} = $null
[System.Data.DataTable]${script:MdbCommand.Schema} = $null
[System.Collections.Generic.SortedDictionary[string,psobject]] `
${script:MdbCommand.Databases} = 
  New-Object "System.Collections.Generic.SortedDictionary[string,psobject]"

#Localized Variable
${script:MdbCommand.LocalizedData} = Data {
  ConvertFrom-StringData @"
    Language = dbLangGeneral
    Constant = ;LANGID=0x0409;CP=1252;COUNTRY=0
    ERROR_OCCURRED = error occurred.
"@
}
Import-LocalizedData -BindingVariable MdbCommand.LocalizedData `
                     -FileName MdbCommandLocalizedData `
                     -ErrorAction SilentlyContinue

#Base Function
function New-MdbException
{
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$true, Position=0)]
    [System.Management.Automation.InvocationInfo]
    $InvocationInfo,
    [System.Management.Automation.Errorrecord]
    $ErrorRecord,
    [string]
    $Message,
    [switch]
    $Out
  )
  $MdbError = New-Object psobject `
  | select InvocationInfo, ErrorRecord, Message, Out
  $MdbError.InvocationInfo = $InvocationInfo
  $MdbError.ErrorRecord = $ErrorRecord
  $MdbError.Message = $Message
  $MdbError.Out = [bool]$Out
  $MdbError.psobject.TypeNames.Insert(0, "MdbCommand.MdbError")
  $global:ChildMdbError = $null
  if ($ErrorRecord -and
      $ErrorRecord.TargetObject -and
      $ErrorRecord.TargetObject.psobject.TypeNames -eq "MdbCommand.MdbError")
  {
     $global:ChildMdbError = $ErrorRecord.TargetObject
     $MdbError.Message `
       += "Call => $($ChildMdbError.InvocationInfo.MyCommand.Name)"
  }
  $MdbError
  if ($Out)
  {
    ${script:MdbCommand.LocalizedData}.ERROR_OCCURRED | Out-Host
    if ($ChildMdbError)
    {
      $MdbError, $ChildMdbError | Out-Host
    }
    else
    {
      $MdbError | Out-Host
    }
  }
}
function Get-MdbFilePath
{
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$true, Position=0)]
    [ValidateNotNullOrEmpty()]
    [string]
    $File,
    [switch]
    $ErrorOut
  )
  if (-not (Test-Path -LiteralPath $File -IsValid -PathType Leaf) -or
      (Test-Path -LiteralPath $File -PathType Container) -or
      -not ($File -like "*.mdb"))
  {
    throw New-MdbException $MyInvocation -Message "NG File" -Out:$ErrorOut
  }
  $Parent = Split-Path -Parent -Path $File
  if (-not $Parent)
  {
    $Parent = "$pwd"
  }
  if (-not (Test-Path -LiteralPath $Parent -PathType Container))
  {
    throw New-MdbException $MyInvocation -Message "NG Folder" -Out:$ErrorOut
  }
  [string](Join-Path (Convert-Path $Parent) (Split-Path -Leaf -Path $File))
}
function New-MdbFile
{
# .ExternalHelp ./MdbCommand-Help.xml
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$true, Position=0)]
    [ValidateNotNullOrEmpty()]
    [string]
    $File,
    [parameter(Mandatory=$false, Position=1)]
    [AllowNull()]
    [string]
    $Password,
    [parameter(Mandatory=$false, Position=2)]
    [ValidateNotNullOrEmpty()]
    [string]
    $Locale,
    [switch]
    $ErrorOut
  )
  ${script:MdbCommand.FilePath} = $null
  try
  {
    ${script:MdbCommand.FilePath} = Get-MdbFilePath -File $File
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_ -Out:$ErrorOut
  }
  $ComEngine = $null
  $ComDatabase = $null
  try
  {
    if (-not $Locale)
    {
      $Locale = ${script:MdbCommand.LocalizedData}.Constant
    }
    $Option = 0x40 # dbVersion40
    if ($Password)
    {
      $Locale += ";pwd=$Password"
      $Option = $Option -bor 0x02 # dbEncrypt
    }
    $ComEngine = New-Object -ComObject Dao.DbEngine.36
    $ComDatabase = $ComEngine.CreateDatabase(
      ${script:MdbCommand.FilePath}, $Locale, $Option)
    $ComDatabase.Close()
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_ -Out:$ErrorOut
  }
  finally
  {
    if ($ComDatabase)
    {
      [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComDatabase) `
      | Out-Null
    }
    if ($ComEngine)
    {
      [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComEngine) `
      | Out-Null
    }
  }
}
function New-MdbConnection
{
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$true, Position=0)]
    [ValidateNotNullOrEmpty()]
    [string]
    $File,
    [parameter(Mandatory=$false, Position=1)]
    [AllowNull()]
    [string]
    $Password
  )
  ${script:MdbCommand.FilePath} = $null
  try
  {
    ${script:MdbCommand.FilePath} = Get-MdbFilePath -File $File
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_
  }
  ${script:MdbCommand.Connection} = $null
  try
  {
    $Builder = New-Object System.Data.OleDb.OleDbConnectionStringBuilder
    $Builder.Provider = "Microsoft.Jet.OLEDB.4.0"
    $Builder["Data Source"] = ${script:MdbCommand.FilePath}
    if ($Password)
    {
      $Builder["Jet OLEDB:Database Password"]= $Password
    }
    $Builder["Jet OLEDB:Database Locking Mode"] = "1"
    ${script:MdbCommand.Connection} =
      New-Object System.Data.OleDb.OleDbConnection($Builder.ConnectionString)
    ${script:MdbCommand.Connection}.Open()
    ${script:MdbCommand.Connection}
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_
  }
}
function New-MdbTransaction
{
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$true,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [System.Data.OleDb.OleDbConnection]
    $Connection
  )
  ${script:MdbCommand.Transaction} = $null
  try
  {
    ${script:MdbCommand.Transaction} = $Connection.BeginTransaction()
    ${script:MdbCommand.Transaction} 
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_
  }
}
function New-MdbCommand
{
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$true,
               ValueFromPipeline=$true,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [System.Data.OleDb.OleDbConnection]
    $Connection,
    [parameter(Mandatory=$false,
               ValueFromPipelineByPropertyName=$true)]
    [AllowNull()]
    [System.Data.OleDb.OleDbTransaction]
    $Transaction,
    [parameter(Mandatory=$true, Position=0)]
    [ValidateNotNullOrEmpty()]
    [Alias("SQL")]
    [string]
    $CommandText
  )
  ${script:MdbCommand.Command} = $null
  try
  {
    ${script:MdbCommand.Command} = $Connection.CreateCommand()
    ${script:MdbCommand.Command}.Transaction = $Transaction
    ${script:MdbCommand.Command}.CommandText = $CommandText
    ${script:MdbCommand.Command}
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_
  }
}
function Set-MdbParameter
{
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$true, Position=0,
               ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [System.Data.OleDb.OleDbCommand]
    $Command,
    [parameter(Mandatory=$true, Position=1)]
    [object[]]
    $Value
  )
  try
  {
    $Count = $Command.Parameters.Count
    if ($Count -gt 0)
    {
      for ($i = 0; $i -lt $Count; $i++)
      {
        $Parameter = $Command.Parameters.Item($i)
        $Parameter.ResetDbType()
        $Parameter.Value = $Value[$i]
      }
    }
    else
    {
      for ($i = 0; $i -lt $Value.Count; $i++)
      {
        $Parameter = $Command.CreateParameter()
        $Parameter.Value = $Value[$i]
        $Command.Parameters.Add($Parameter) | Out-Null
      }
    }
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_
  }
}
function Invoke-MdbCommand
{
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$true, Position=0,
               ValueFromPipeline=$true)]
    [ValidateNotNullOrEmpty()]
    [System.Data.OleDb.OleDbCommand]
    $Command,
    [parameter(Mandatory=$false, Position=1)]
    [string]
    $TypeName = "MdbCommand.MdbRecord"
  )
  ${script:MdbCommand.RecordsAffected} = $null
  ${script:MdbCommand.Schema} = $null
  [System.Data.OleDb.OleDbDataReader]$Reader = $null
  try
  {
    $Reader = $Command.ExecuteReader()
    $FieldCount = $Reader.FieldCount
    if ($FieldCount)
    {
      ${script:MdbCommand.Schema} = $Reader.GetSchemaTable()
      $Names = @(${script:MdbCommand.Schema} | % {$_.ColumnName})
      $Values = New-Object object[] -ArgumentList $FieldCount
    }
    while ($Reader.Read())
    {
      $Count = $Reader.GetValues($Values)
      $Object = New-Object psobject | select $Names
      for ($i = 0; $i -lt $Count; $i++)
      {
        $Object."$($Names[$i])" = $Values[$i]
      }
      $Object.psobject.TypeNames.Insert(0, $TypeName)
      $Object
    }
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_
  }
  finally
  {
    if ($Reader)
    {
      $Reader.Close()
      ${script:MdbCommand.RecordsAffected} = $Reader.RecordsAffected
    }
  }
}
function Get-MdbSchema
{
  [CmdletBinding(DefaultParameterSetName="Table")]
  param(
    [parameter(Mandatory=$true,
               ValueFromPipeline=$true,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [System.Data.OleDb.OleDbConnection]
    $Connection,
    [parameter(ParameterSetName="Table")]
    [parameter(ParameterSetName="Column")]
    [parameter(ParameterSetName="Index")]
    [AllowNull()]
    [string]
    $TableName,
    [parameter(ParameterSetName="Column")]
    [parameter(ParameterSetName="Index")]
    [AllowNull()]
    [string]
    $ColumnName,
    [parameter(ParameterSetName="Index")]
    [AllowNull()]
    [string]
    $IndexName,
    [parameter(ParameterSetName="Query")]
    [AllowNull()]
    [string]
    $QueryName,
    [parameter(ParameterSetName="Table")]
    [switch]
    $Table,
    [parameter(ParameterSetName="Column")]
    [switch]
    $Column,
    [parameter(ParameterSetName="Index")]
    [switch]
    $Index,
    [parameter(ParameterSetName="Query")]
    [switch]
    $Query
  )
  function Quote
  {
    param([string]$s)
    "'$(if ($s.IndexOf("'")){$s.Replace("'", "''")} else {$s})'"
  }
  [string]$Filter = $null
  [string]$Sort = $null
  [string]$NameFilter = $null
  try
  {
    switch ($PsCmdlet.ParameterSetName)
    {
      Table
      {
        $Filter = "TABLE_TYPE Not In ('ACCESS TABLE', 'SYSTEM TABLE', 'VIEW')"
        $Sort = "TABLE_NAME"
        if ($TableName)
        {
          $Filter += " And TABLE_NAME Like $(Quote $TableName)"
        }
        $Connection.GetSchema("Tables").Select($Filter, $Sort) `
        | select @(
          @{l="TableName"; e="TABLE_NAME"}
          @{l="TableType"; e="TABLE_TYPE"}
          @{l="Description"; e="DESCRIPTION"}) `
        | % {$_.psobject.TypeNames.Insert(0, "MdbCommand.MdbTable"); $_}
      }
      Column
      {
        $DataTypeConverter = {
          switch($_)
          {
            {$_.DATA_TYPE -eq   2}{"SMALLINT" ; break}
            {$_.DATA_TYPE -eq   3}{"INTEGER" ; break}
            {$_.DATA_TYPE -eq   4}{"REAL" ; break}
            {$_.DATA_TYPE -eq   5}{"FLOAT" ; break}
            {$_.DATA_TYPE -eq   6}{"MONEY" ; break}
            {$_.DATA_TYPE -eq   7}{"DATETIME" ; break}
            {$_.DATA_TYPE -eq  11}{"BIT" ; break}
            {$_.DATA_TYPE -eq  17}{"TINYINT" ; break}
            {$_.DATA_TYPE -eq  72}{"UNIQUEIDENTIFER" ; break}
            {$_.DATA_TYPE -eq 128 -and
             $_.CHARACTER_MAXIMUM_LENGTH -eq 0}{"IMAGE" ; break}
            {$_.DATA_TYPE -eq 128 -and
             $_.CHARACTER_MAXIMUM_LENGTH -ne 0}{"BINARY" ; break}
            {$_.DATA_TYPE -eq 130 -and
             $_.CHARACTER_MAXIMUM_LENGTH -eq 0}{"TEXT" ; break}
            {$_.DATA_TYPE -eq 130 -and
             $_.CHARACTER_MAXIMUM_LENGTH -ne 0 -and
             $_.COLUMN_FLAGS -eq 122}{"CHARACTER" ; break}
            {$_.DATA_TYPE -eq 130 -and
             $_.CHARACTER_MAXIMUM_LENGTH -ne 0 -and
             $_.COLUMN_FLAGS -eq 106}{"VARCHAR" ; break}
            {$_.DATA_TYPE -eq 131}{"DECIMAL" ; break}
            default {"unknown"}
          }
        }
        $AliasDataTypeConverter = {
          switch($_)
          {
            {$_.DATA_TYPE -eq   2}{"short" ; break}
            {$_.DATA_TYPE -eq   3}{"int" ; break}
            {$_.DATA_TYPE -eq   4}{"single" ; break}
            {$_.DATA_TYPE -eq   5}{"double" ; break}
            {$_.DATA_TYPE -eq   6}{"currency" ; break}
            {$_.DATA_TYPE -eq   7}{"date" ; break}
            {$_.DATA_TYPE -eq  11}{"yesno" ; break}
            {$_.DATA_TYPE -eq  17}{"byte" ; break}
            {$_.DATA_TYPE -eq  72}{"guid" ; break}
            {$_.DATA_TYPE -eq 128 -and 
             $_.CHARACTER_MAXIMUM_LENGTH -eq 0}{"oleobject" ; break}
            {$_.DATA_TYPE -eq 128 -and
             $_.CHARACTER_MAXIMUM_LENGTH -ne 0}{"binary" ; break}
            {$_.DATA_TYPE -eq 130 -and
             $_.CHARACTER_MAXIMUM_LENGTH -eq 0}{"memo" ; break}
            {$_.DATA_TYPE -eq 130 -and
             $_.CHARACTER_MAXIMUM_LENGTH -ne 0 -and
             $_.COLUMN_FLAGS -eq 122}{"char" ; break}
            {$_.DATA_TYPE -eq 130 -and
             $_.CHARACTER_MAXIMUM_LENGTH -ne 0 -and
             $_.COLUMN_FLAGS -eq 106}{"string" ; break}
            {$_.DATA_TYPE -eq 131}{"numeric"; break}
            default {"unknown"}
          }
        }
        $SizeConverter = {
          switch($_)
          {
            {$_.DATA_TYPE -eq 130 -and
             $_.CHARACTER_MAXIMUM_LENGTH -ne 0}
            {
              $_.CHARACTER_MAXIMUM_LENGTH
              break
            }
            {$_.DATA_TYPE -eq 131}
            {
              "$($_.NUMERIC_PRECISION),$($_.NUMERIC_SCALE)"
              break
            }
            default {$null}
          }
        }
        $PrimaryKeyConverter = {
          $PkGuid = [System.Data.OleDb.OleDbSchemaGuid]::Primary_Keys
          $Restrictions = @($null, $null, $_.TABLE_NAME)
          $Filter = "COLUMN_NAME = $(Quote $_.COLUMN_NAME)"
          if ($Connection.GetOleDbSchemaTable(`
              $PkGuid, $Restrictions).Select($Filter))
          {
            "YES"
          }
          else
          {
            $null
          }
        }
        $TypeName = 
          "System.Collections.Generic.Dictionary[string,System.Data.DataTable]"
        $TableSchemas = New-Object $TypeName
        $AutoIncrementConverter = {
          if ($TableSchemas.ContainsKey($_.TABLE_NAME))
          {
            $DataTable = $TableSchemas.Item($_.TABLE_NAME)
          }
          else
          {
            $SQL = "Select * From [$($_.TABLE_NAME)]"
            $Command = $null
            $Adapter = $null
            try
            {
              $Command = $Connection | New-MdbCommand -CommandText $SQL
              $Adapter = New-Object System.Data.OleDb.OleDbDataAdapter($Command)
              $Adapter.MissingSchemaAction = 
                [System.Data.MissingSchemaAction]::AddWithKey
              $DataTable = New-Object System.Data.DataTable
              $Adapter.FillSchema($DataTable, [System.Data.SchemaType]::Source)
              $TableSchemas.Add($_.TABLE_NAME, $DataTable)
            }
            catch
            {
              "ERROR"
            }
            finally
            {
              if ($Adapter)
              {
                $Adapter.Dispose()
              }
              if ($Command)
              {
                $Command.Dispose()
              }
            }
          }
          if ($DataTable.Columns.Item($_.COLUMN_NAME).AutoIncrement)
          {
            "YES"
          }
          else
          {
            $Null
          }
        }
        $Filter = "TABLE_TYPE In ('ACCESS TABLE', 'SYSTEM TABLE', 'VIEW')"
        $SystemTables = $Connection.GetSchema("Tables").Select($Filter) `
                        | % {Quote $_.TABLE_NAME}
        $OFS = ","
        $Filter = "TABLE_NAME Not In ($SystemTables)"
        $Sort = "TABLE_NAME, ORDINAL_POSITION"
        if ($TableName)
        {
          $Filter += " And TABLE_NAME Like $(Quote $TableName)"
        }
        if ($ColumnName)
        {
          $Filter += " And COLUMN_NAME Like $(Quote $ColumnName)"
        }
        $Connection.GetSchema("Columns").Select($Filter, $Sort) `
        | select @(
          @{l="TableName"; e="TABLE_NAME"}
          @{l="Position"; e="ORDINAL_POSITION"}
          @{l="ColumnName"; e="COLUMN_NAME"}
          @{l="DataType"; e={& $DataTypeConverter $_}}
          @{l="AliasType"; e={& $AliasDataTypeConverter $_}}
          @{l="Size"; e={& $SizeConverter $_}}
          @{l="NotNull"; e={if($_.IS_NULLABLE) {$null} else {"YES"}}}
          @{l="Default"; e="COLUMN_DEFAULT"}
          @{l="PrimaryKey"; e={& $PrimaryKeyConverter $_}}
          @{l="AutoIncrement"; e={& $AutoIncrementConverter $_}}
          @{l="Description"; e="DESCRIPTION"}) `
        | % {$_.psobject.TypeNames.Insert(0, "MdbCommand.MdbColumn"); $_} 
        if ($TableSchemas.Count)
        {
          foreach ($TableSchema in $TableSchemas.GetEnumerator())
          {
            $TableSchema.Value.Dispose()
          }
        }
      }
      Index
      {
        $Filter = "TABLE_TYPE In ('ACCESS TABLE', 'SYSTEM TABLE')"
        $SystemTables = $Connection.GetSchema("Tables").Select($Filter) `
                        | % {Quote $_.TABLE_NAME}
        $OFS = ","
        $Filter = "TABLE_NAME Not In ($SystemTables)"
        $Sort = "TABLE_NAME, INDEX_NAME, ORDINAL_POSITION"
        if ($TableName)
        {
          $Filter += " And TABLE_NAME Like $(Quote $TableName)"
        }
        if ($ColumnName)
        {
          $Filter += " And COLUMN_NAME Like $(Quote $ColumnName)"
        }
        if ($IndexName)
        {
          $Filter += " And INDEX_NAME Like $(Quote $IndexName)"
        }
        $Connection.GetSchema("Indexes").Select($Filter, $Sort) `
        | select @(
          @{l="TableName"; e="TABLE_NAME"}
          @{l="IndexName"; e="INDEX_NAME"}
          @{l="Primary"; e={$(if ($_.PRIMARY_KEY) {"YES"} else {$null})}}
          @{l="Unique"; e={$(if ($_.UNIQUE) {"YES"} else {$null})}}
          @{l="Position"; e="ORDINAL_POSITION"}
          @{l="ColumnName"; e="COLUMN_NAME"}
          @{l="Collation";  e={if ($_.COLLATION -eq 1) {"ASC" }
                               elseif ($_.COLLATION -eq 2) {"DESC"}}}) `
          | % {$_.psobject.TypeNames.Insert(0, "MdbCommand.MdbIndex"); $_} 
      }
      Query
      {
        $Filter = ""
        if ($QueryName)
        {
          $Filter += "PROCEDURE_NAME Like $(Quote $QueryName)"
        }
        $Querys = @(
          $Connection.GetSchema("Procedures").Select($Filter) `
          | select @(
            @{l="QueryName"; e="PROCEDURE_NAME"}
            @{l="QueryType"; e={"ACTION"}}
            @{l="SQL"; e="PROCEDURE_DEFINITION"})
        )
        $Filter = ""
        if ($QueryName)
        {
          $Filter += "TABLE_NAME Like $(Quote $QueryName)"
        }
        $Querys += @(
          $Connection.GetSchema("Views").Select($Filter) `
          | select @(
            @{l="QueryName"; e="TABLE_NAME"}
            @{l="QueryType"; e={"VIEW"}}
            @{l="SQL"; e="VIEW_DEFINITION"})
        )
        $Querys | sort QueryName `
        | % {$_.psobject.TypeNames.Insert(0, "MdbCommand.MdbQuery"); $_} 
      }
    }
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_
  }
}
#Automation Bridge Function
function Add-Mdb
{
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$true,Position=0,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DatabaseName,
    [parameter(Mandatory=$true,Position=1,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [System.Data.OleDb.OledbConnection]
    $Connection,
    [parameter(Mandatory=$true,Position=2,
               ValueFromPipelineByPropertyName=$true)]
    [AllowNull()]
    [System.Data.OleDb.OleDbTransaction]
    $Transaction,
    [parameter(Mandatory=$true,Position=3,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $File
  )
  try
  {
    $Mdb = New-Module -AsCustomObject `
      -ArgumentList $DatabaseName, $Connection, $Transaction, $File `
      -ScriptBlock `
      {
        param(
          [string]
          $DatabaseName,
          [System.Data.OleDb.OledbConnection]
          $Connection,
          [System.Data.OleDb.OleDbTransaction]
          $Transaction,
          [string]
          $File
        )
        Export-ModuleMember -Variable *
      }
    $Mdb.psobject.TypeNames.Insert(0, "MdbCommand.Mdb")
    ${script:MdbCommand.Databases}.add($DatabaseName, $Mdb)
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_
  }
}
function Get-MdbAvailable
{
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$false, Position=0,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DatabaseName
  )
  if ($DatabaseName)
  {
    if (${script:MdbCommand.Databases}.ContainsKey($DatabaseName))
    {
      ${script:MdbCommand.Databases}.$DatabaseName
    }
    else
    {
      throw New-MdbException $MyInvocation -Message "NG DatabaseName"
    }
  }
  else
  {
    switch (${script:MdbCommand.Databases}.Count)
    {
      0
      {
        throw New-MdbException $MyInvocation -Message "No Mdb"
        break
      }
      1
      {
        ${script:MdbCommand.Databases}.Values
        break
      }
      default
      {
        throw New-MdbException $MyInvocation -Message "Witch Mdb?"
      }
    }
  }
}
function Remove-Mdb
{
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$true,Position=0,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DatabaseName
  )
  try
  {
    ${script:MdbCommand.Databases}.Remove($DatabaseName) | Out-Null
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_
  }
}
#Automation Function
function Get-Mdb
{
# .ExternalHelp ./MdbCommand-Help.xml
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$false, Position=0,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DatabaseName
  )
  try
  {
    if ($DatabaseName)
    {
      if (${script:MdbCommand.Databases}.ContainsKey($DatabaseName))
      {
        ${script:MdbCommand.Databases}.$DatabaseName
      }
    }
    else
    {
      ${script:MdbCommand.Databases}.Values
    }
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_
  }
}
function Open-Mdb
{
# .ExternalHelp ./MdbCommand-Help.xml
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$false,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DatabaseName = "main",
    [parameter(Mandatory=$true, Position=0,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $File,
    [parameter(Mandatory=$false, Position=1,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $Password,
    [switch]
    $Create,
    [switch]
    $PassThru
  )
  [string]$FilePath = $null
  if (Get-Mdb -DatabaseName $DatabaseName)
  {
    throw New-MdbException $MyInvocation -Message "NG DatabaseName" -Out
  }
  try
  {
    $FilePath = Get-MdbFilePath -File $File
    if ($Create)
    {
      New-MdbFile -File $FilePath -Password $Password
    }
    $Connection = New-MdbConnection -File $FilePath -Password $Password
    Add-Mdb -DatabaseName $DatabaseName -Connection $Connection `
            -Transaction $null -File $FilePath
    if ($PassThru)
    {
      Get-MdbAvailable -DatabaseName $DatabaseName
    }
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_ -Out
  }
}
function Close-Mdb
{
# .ExternalHelp ./MdbCommand-Help.xml
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$false, Position=0,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DatabaseName
  )
  $Mdb = $null
  try
  {
    $ht = @{}
    if ($DatabaseName)
    {
      $ht.Add("DatabaseName", $DatabaseName)
    }
    $Mdb = Get-MdbAvailable @ht
    $Mdb.Connection.Close()
    $Mdb | Remove-Mdb
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_ -Out
  }
}
function Invoke-Mdb
{
# .ExternalHelp ./MdbCommand-Help.xml
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$false,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DatabaseName,
    [parameter(Mandatory=$true, Position=0,
               ValueFromPipeline=$true,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $SQL,
    [parameter(Mandatory=$false, Position=1,
               ValueFromPipelineByPropertyName=$true)]
    [Object[]]
    $Value
  )
  $Mdb = $null
  [System.Data.OleDb.OleDbCommand]$Command = $null
  try
  {
    $ht = @{}
    if ($DatabaseName)
    {
      $ht.Add("DatabaseName", $DatabaseName)
    }
    $Mdb = Get-MdbAvailable @ht
    $Command = $Mdb | New-MdbCommand -CommandText $SQL
    if ($Value)
    {
      $Command | Set-MdbParameter -Value $Value
    }
    $Command | Invoke-MdbCommand
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_ -Out
  }
  finally
  {
    if ($Command)
    {
      $Command.Dispose()
    }
  }
}
function Invoke-MdbBatch
{
# .ExternalHelp ./MdbCommand-Help.xml
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$false,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DatabaseName,
    [parameter(Mandatory=$true, Position=0,
               ValueFromPipeline=$true,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $SQL,
    [parameter(Mandatory=$false, Position=1,
               ValueFromPipelineByPropertyName=$true)]
    [System.Collections.ArrayList]
    $ValueList
  )
  $Mdb = $null
  [System.Data.OleDb.OleDbCommand]$Command = $null
  ${script:MdbCommand.BatchTotalRecordsAffected} = 0
  try
  {
    $ht = @{}
    if ($DatabaseName)
    {
      $ht.Add("DatabaseName", $DatabaseName)
    }
    $Mdb = Get-MdbAvailable @ht
    $Command = $Mdb | New-MdbCommand -CommandText $SQL
    if ($ValueList -and $ValueList.Count)
    {
      foreach ($Value in $ValueList.GetEnumerator())
      {
        $Command | Set-MdbParameter -Value $Value
        $Command | Invoke-MdbCommand
        ${script:MdbCommand.BatchTotalRecordsAffected} `
           += ${script:MdbCommand.RecordsAffected}
      }
    }
    else
    {
      $Command | Invoke-MdbCommand
      ${script:MdbCommand.BatchTotalRecordsAffected} `
         += ${script:MdbCommand.RecordsAffected}
    }
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_ -Out
  }
  finally
  {
    if ($Command)
    {
      $Command.Dispose()
    }
  }
}
function Start-MdbTransaction
{
# .ExternalHelp ./MdbCommand-Help.xml
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$false, Position=0,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DatabaseName
  )
  $Mdb = $null
  try
  {
    $ht = @{}
    if ($DatabaseName)
    {
      $ht.Add("DatabaseName", $DatabaseName)
    }
    $Mdb = Get-MdbAvailable @ht
    $Transaction = $Mdb | New-MdbTransaction
    $Mdb.Transaction = $Transaction
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_ -Out
  }
}
function Complete-MdbTransaction
{
# .ExternalHelp ./MdbCommand-Help.xml
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$false, Position=0,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DatabaseName
  )
  $Mdb = $null
  try
  {
    $ht = @{}
    if ($DatabaseName)
    {
      $ht.Add("DatabaseName", $DatabaseName)
    }
    $Mdb = Get-MdbAvailable @ht
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_ -Out
  }
  if (-not $Mdb.Transaction)
  {
    throw New-MdbException $MyInvocation -Message "NO Transaction" -Out
  }
  try
  {
    $Mdb.Transaction.Commit()
    $Mdb.Transaction = $null
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_ -Out
  }
}
function Undo-MdbTransaction
{
# .ExternalHelp ./MdbCommand-Help.xml
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$false, Position=0,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DatabaseName
  )
  $Mdb = $null
  try
  {
    $ht = @{}
    if ($DatabaseName)
    {
      $ht.Add("DatabaseName", $DatabaseName)
    }
    $Mdb = Get-MdbAvailable @ht
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_ -Out
  }
  if (-not $Mdb.Transaction)
  {
    throw New-MdbException $MyInvocation -Message "NO Transaction" -Out
  }
  try
  {
    $Mdb.Transaction.Rollback()
    $Mdb.Transaction = $null
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_ -Out
  }
}
#Information
function Get-MdbTable
{
# .ExternalHelp ./MdbCommand-Help.xml
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$false,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DatabaseName,
    [parameter(Mandatory=$false, Position=0,
               ValueFromPipelineByPropertyName=$true)]
    [string]
    $TableName
  )
  $Mdb = $null
  try
  {
    $ht = @{}
    if ($DatabaseName)
    {
      $ht.Add("DatabaseName", $DatabaseName)
    }
    $Mdb = Get-MdbAvailable @ht
    $Mdb | Get-MdbSchema -Table -TableName $TableName
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_ -Out
  }
}
function Get-MdbColumn
{
# .ExternalHelp ./MdbCommand-Help.xml
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$false,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DatabaseName,
    [parameter(Mandatory=$false, Position=0,
               ValueFromPipelineByPropertyName=$true)]
    [AllowNull()]
    [string]
    $TableName,
    [parameter(Mandatory=$false, Position=1,
               ValueFromPipelineByPropertyName=$true)]
    [AllowNull()]
    [string]
    $ColumnName
  )
  $Mdb = $null
  try
  {
    $ht = @{}
    if ($DatabaseName)
    {
      $ht.Add("DatabaseName", $DatabaseName)
    }
    $Mdb = Get-MdbAvailable @ht
    $Mdb | Get-MdbSchema -Column -TableName $TableName -ColumnName $ColumnName
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_ -Out
  }
}
function Get-MdbIndex
{
# .ExternalHelp ./MdbCommand-Help.xml
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$false,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DatabaseName,
    [parameter(Mandatory=$false, Position=0,
               ValueFromPipelineByPropertyName=$true)]
    [AllowNull()]
    [string]
    $TableName,
    [parameter(Mandatory=$false, Position=1,
               ValueFromPipelineByPropertyName=$true)]
    [AllowNull()]
    [string]
    $ColumnName,
    [parameter(Mandatory=$false, Position=2,
               ValueFromPipelineByPropertyName=$true)]
    [AllowNull()]
    [string]
    $IndexName
  )
  $Mdb = $null
  try
  {
    $ht = @{}
    if ($DatabaseName)
    {
      $ht.Add("DatabaseName", $DatabaseName)
    }
    $Mdb = Get-MdbAvailable @ht
    $Mdb | Get-MdbSchema -Index -TableName $TableName `
                         -ColumnName $ColumnName -IndexName $IndexName
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_ -Out
  }
}
function Get-MdbQuery
{
# .ExternalHelp ./MdbCommand-Help.xml
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$false,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DatabaseName,
    [parameter(Mandatory=$false, Position=0,
               ValueFromPipelineByPropertyName=$true)]
    [AllowNull()]
    [string]
    $QueryName
  )
  $Mdb = $null
  try
  {
    $ht = @{}
    if ($DatabaseName)
    {
      $ht.Add("DatabaseName", $DatabaseName)
    }
    $Mdb = Get-MdbAvailable @ht
    $Mdb | Get-MdbSchema -Query -QueryName $QueryName
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_ -Out
  }
}
function Add-MdbQuery
{
# .ExternalHelp ./MdbCommand-Help.xml
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$false,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DatabaseName,
    [parameter(Mandatory=$true, Position=0,
               ValueFromPipelineByPropertyName=$true)]
    [AllowNull()]
    [string]
    $QueryName,
    [parameter(Mandatory=$true, Position=1,
               ValueFromPipelineByPropertyName=$true)]
    [AllowNull()]
    [string]
    $SQL
  )
  $Mdb = $null
  [System.Data.OleDb.OleDbCommand]$Command = $null
  try
  {
    $ht = @{}
    if ($DatabaseName)
    {
      $ht.Add("DatabaseName", $DatabaseName)
    }
    $Mdb = Get-MdbAvailable @ht
    $CommandText = "Create Procedure [$QueryName] As $SQL"
    $Command = $Mdb | New-MdbCommand -CommandText $CommandText
    $Command | Invoke-MdbCommand
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_ -Out
  }
  finally
  {
    if ($Command)
    {
      $Command = $null
    }
  }
}
function Invoke-MdbQuery
{
# .ExternalHelp ./MdbCommand-Help.xml
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$false,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DatabaseName,
    [parameter(Mandatory=$true, Position=0,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $QueryName,
    [parameter(Mandatory=$false, Position=1,
               ValueFromPipelineByPropertyName=$true)]
    [Object[]]
    $Value
  )
  $Mdb = $null
  $Query = $null
  try
  {
    $ht = @{}
    if ($DatabaseName)
    {
      $ht.Add("DatabaseName", $DatabaseName)
    }
    $Mdb = Get-MdbAvailable @ht
    $Query = Get-MdbQuery -QueryName $QueryName
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_ -Out
  }
  if (-not $Query)
  {
    throw New-MdbException $MyInvocation -Message "NG QueryName" -Out
  }
  [System.Data.OleDb.OleDbCommand]$Command = $null
  try
  {
    $Command = $Mdb | New-MdbCommand -CommandText $Query.SQL
    if ($Value)
    {
      $Command | Set-MdbParameter -Value $Value
    }
    $Command | Invoke-MdbCommand
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_ -Out
  }
  finally
  {
    if ($Command)
    {
      $Command.Dispose()
    }
  }
}
function Remove-MdbQuery
{
# .ExternalHelp ./MdbCommand-Help.xml
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$false,
               ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $DatabaseName,
    [parameter(Mandatory=$true, Position=0,
               ValueFromPipelineByPropertyName=$true)]
    [AllowNull()]
    [string]
    $QueryName
  )
  $Mdb = $null
  $Command = $null
  try
  {
    $ht = @{}
    if ($DatabaseName)
    {
      $ht.Add("DatabaseName", $DatabaseName)
    }
    $Mdb = Get-MdbAvailable @ht
    $CommandText = "Drop Procedure [$QueryName]"
    $Command = $Mdb | New-MdbCommand -CommandText $CommandText
    $Command | Invoke-MdbCommand
  }
  catch
  {
    throw New-MdbException $MyInvocation -ErrorRecord $_ -Out
  }
  finally
  {
    if ($Command)
    {
      $Command = $null
    }
  }
}

#Inport Console
. $psScriptRoot\MdbCommandConsole.ps1
#Inport Alias Setting
. $psScriptRoot\MdbCommandAlias.ps1

#Export Member
Export-ModuleMember -Variable *MdbCommand* -Function *Mdb*

#Closing
$MyInvocation.MyCommand.ScriptBlock.Module.OnRemove = {
  try
  {
    ${script:MdbCommand.Databases}.Values | % {$_.Connection.Close()}
  }
  catch
  {
    Write-Warning "$_"
  }
  "bye" | Out-Host
}
