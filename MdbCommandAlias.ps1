#requires -version 2.0

#MdbCommand : Alias Setting
#Author rancomma@gmail.com twitter @rancomma

[hashtable]${script:MdbCommand.AliasConfig} = @{
"close"        = "Close-Mdb"
"quit"         = "Close-Mdb"
"commit"       = "Complete-MdbTransaction"
"databases"    = "Get-Mdb"
"columns"      = "Get-MdbColumn"
"path"         = "Get-MdbFilePath"
"indices"      = "Get-MdbIndex"
"querys"       = "Get-MdbQuery"
"tables"       = "Get-MdbTable"
"execute"      = "Invoke-Mdb"
"executebatch" = "Invoke-MdbBatch"
"executequery" =  "Invoke-MdbQuery"
"create"       = "New-MdbFile"
"open"         = "Open-Mdb"
"console"      = "Start-MdbConsole"
"begin"        = "Start-MdbTransaction"
"rollback"     = "Undo-MdbTransaction"
}

function Use-MdbAliasCommand
{
# .ExternalHelp ./MdbCommand-Help.xml
  param(
    [string]
    $Prefix = "",
    [string]
    $Postfix = ""
  )
  foreach ($ht in ${script:MdbCommand.AliasConfig}.GetEnumerator())
  {
    $Name = "$Prefix$($ht.Key)$Postfix"
    New-Alias $Name -Value $ht.Value -Scope global -Force
  }
}
function Use-MdbDotCommand
{
# .ExternalHelp ./MdbCommand-Help.xml
  param(
    [string]
    $Prefix = ".",
    [string]
    $Postfix = ""
  )
  Use-MdbAliasCommand -PreFix $PreFix -PostFix $Postfix
}
function Open-MdbWithDotCommand
{
# .ExternalHelp ./MdbCommand-Help.xml
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
    $PassThru,
    [String]
    $Prefix = ".",
    [String]
    $Postfix = ""
  )
  $ht = @{}
  if ($DatabaseName) {$ht.DatabaseName = $DatabaseName}
  if ($File) {$ht.File= $File}
  if ($Password) {$ht.Password = $Password}
  if ($Create) {$ht.Create = $Create}
  if ($PassThru) {$ht.PassThru = $PassThru}
  Open-Mdb @ht
  Use-MdbAliasCommand -PreFix $PreFix -PostFix $Postfix
}
New-Alias ".mdb" -Value Open-MdbWithDotCommand -Scope global -Force
