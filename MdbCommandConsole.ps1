#requires -version 2.0

#MdbCommand : Console
#Author rancomma@gmail.com twitter @rancomma

#Console Config
[bool]$script:AutoSize = $true
[bool]$script:Echo = $false
[bool]$script:Headers = $true
[bool]$script:Wrap = $true
[string]$script:NullValue = $null
[int[]]$script:Width = @()
[string]$script:DefaultPrompt = "mdb> "
[string]$script:NestedPrompt  = "...> "

#DotCommand Config
[hashtable]$script:DotCommandConfig = @{
  ".autosize"   = @{cnt = 1;   op = "eq"; onoff = $true}
  ".begin"      = @{cnt = 0;   op = "eq"; onoff = $false}
  ".columns"    = @{cnt = 2;   op = "le"; onoff = $false}
  ".commit"     = @{cnt = 0;   op = "eq"; onoff = $false}
  ".databases"  = @{cnt = 0;   op = "eq"; onoff = $false}
  ".echo"       = @{cnt = 1;   op = "eq"; onoff = $true}
  ".exit"       = @{cnt = 0;   op = "eq"; onoff = $false}
  ".headers"    = @{cnt = 1;   op = "eq"; onoff = $true}
  ".help"       = @{cnt = 0;   op = "eq"; onoff = $false}
  ".indices"    = @{cnt = 3;   op = "le"; onoff = $false}
  ".nullvalue"  = @{cnt = 1;   op = "eq"; onoff = $false}
  ".quit"       = @{cnt = 0;   op = "eq"; onoff = $false}
  ".rollback"   = @{cnt = 0;   op = "eq"; onoff = $false}
  ".show"       = @{cnt = 0;   op = "eq"; onoff = $false}
  ".tables"     = @{cnt = 1;   op = "le"; onoff = $false}
  ".views"      = @{cnt = 1;   op = "le"; onoff = $false}
  ".width"      = @{cnt = 255; op = "le"; onoff = $false}
  ".wrap"       = @{cnt = 1;   op = "eq"; onoff = $true}
}
[string[]]$script:DotCommandList = $DotCommandConfig.Keys

[string]$script:HelpMessage = @"
.autosize ON|OFF       Like Format-Table AutoSize parameter
.begin                 Start transaction
.columns ?TABLE?       List columns
.commit                Commit transaction
.databases             Show database Info
.echo ON|OFF           Turn command echo on or off
.exit                  Exit this console
.headers ON|OFF        Turn display of headers on or off
.help                  Show this message
.indices ?T? ?C? ?I?   List indices
.nullvalue STRING      Print STRING in place of NULL values
.quit                  Exit this console
.rollback              Rollback transaction
.show                  Show the current values for various settings
.tables ?TABLE?        List tables
.views ?VIEW?          List querys and procedures
.width NUM1 NUM2 ...   Set column widths (If ".width" Then clear widths)
.wrap ON|OFF           Like Format-Table Wrap parameter
; or / or GO           Enter SQL statements terminated
"@

#Console
function Start-MdbConsole
{
# .ExternalHelp ./MdbCommand-Help.xml
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$false,
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
    #Variable
    [string]$Line = $null
    [string]$MultiLine = $null
    [bool]$ExitStatus = $false
    Do
    {
      #Loop Variable
      [string]$DotCommand = $null
      [string[]]$ArgumentList = @()
      [bool]$OnOff = $false

      #Output Prompt
      $prompt = $(if (-not $MultiLine) {$DefaultPrompt} else {$NestedPrompt})
      [Console]::Out.Write($prompt)
      #Intput Line
      $Line = [Console]::In.ReadLine()

      #DotCommand Getting
      if (-not $MultiLine -and $Line -and $Line[0] -eq ".")
      {
        $ArgumentList = @($Line.Split(" "))
        $DotCommandMatches = @($DotCommandList -like "$($ArgumentList[0])*")
        if ($DotCommandMatches.Count -eq 1)
        {
          $DotCommand = $DotCommandMatches[0]
        }
        else
        {
          $DotCommand = $ArgumentList[0]
        }
        $Line = $null
      }

      #DotCommand Checking
      if ($DotCommand -and $DotCommandConfig.ContainsKey($DotCommand))
      {
        $cnt = $DotCommandConfig.$DotCommand.cnt
        $op = $DotCommandConfig.$DotCommand.op
        if (-not (iex "$($ArgumentList.Count - 1) -$op $cnt"))
        {
          "error argument count : $DotCommand" | Out-Host
          continue
        }
        if ($DotCommandConfig.$DotCommand.onoff)
        {
          switch ($ArgumentList[1])
          {
            on
            {
              $OnOff = $true
            }
            off
            {
              $OnOff = $false
            }
            default
            {
              "error argument onoff : $DotCommand" | Out-Host
              continue
            }
          }
        }
      }

      #DotCommand Operation
      switch ($DotCommand)
      {
        {-not $DotCommand}
        {
          break
        }
        {"exit", ".quit" -eq $_}
        {
          $ExitStatus = $true
          break
        }
        ".databases"
        {
          $Mdb `
          | Format-Table -Wrap:($script:Wrap) `
                         -AutoSize:($script:AutoSize) `
                         -HideTableHeaders:(-not $script:Headers) `
          | Out-Host
          break
        }
        ".tables"
        {
          $TableName = $ArgumentList[1]
          $Mdb `
          | Get-MdbSchema -Table -TableName  $TableName `
          | Format-Table -Wrap:($script:Wrap) `
                         -AutoSize:($script:AutoSize) `
                         -HideTableHeaders:(-not $script:Headers) `
          | Out-Host
          break
        }
        ".columns"
        {
          $TableName = $ArgumentList[1]
          $CollumnName = $ArgumentList[2]
          $Mdb `
          | Get-MdbSchema -Column -TableName $TableName `
                          -ColumnName $ColumnName `
          | Format-Table -Wrap:($script:Wrap) `
                         -AutoSize:($script:AutoSize) `
                         -HideTableHeaders:(-not $script:Headers) `
          | Out-Host
          break
        }
        ".indices"
        {
          $TableName = $ArgumentList[1]
          $CollumnName = $ArgumentList[2]
          $IndexName = $ArgumentList[3]
          $Mdb `
          | Get-MdbSchema -Index -TableName $TableName `
                          -ColumnName $ColumnName -IndexName $IndexName `
          | Format-Table -Wrap:($script:Wrap) `
                         -AutoSize:($script:AutoSize) `
                         -HideTableHeaders:(-not $script:Headers) `
          | Out-Host
          break
        }
        ".views"
        {
          $QueryName = $ArgumentList[1]
          $Mdb `
          | Get-MdbSchema -Query -QueryName $QueryName `
          | Format-Table -Wrap:($script:Wrap) `
                         -AutoSize:($script:AutoSize) `
                         -HideTableHeaders:(-not $script:Headers) `
          | Out-Host
          break
        }
        ".begin"
        {
          try
          {
            $Transaction = $Mdb | New-MdbTransaction
            $Mdb.Transaction = $Transaction
          }
          catch {
            if ($_.TargetObject.psobject.TypeNames -eq "MdbCommand.MdbError")
            {
              $MdbError = $_.TargetObject
              if ($MdbError.Message)
              {
                $MdbError.Message| Out-Host
              }
              else
              {
                $MdbError.ErrorRecord.Exception.InnerException.Message `
                | Out-Host
              }
            }
            else
            {
              $_.Exception.InnerException.Message | Out-Host
            }
          }
          break
        }
        ".commit"
        {
          try
          {
            if (-not $Mdb.Transaction)
            {
               "no transaction" | Out-Host
            }
            else
            {
              $Mdb.Transaction.Commit()
              $Mdb.Transaction = $null
            }
          }
          catch
          {
            $_.Exception.InnerException.Message | Out-Host
          }
          break
        }
        ".rollback"
        {
          try
          {
            if (-not $Mdb.Transaction)
            {
               "no transaction" | Out-Host
            }
            else
            {
              $Mdb.Transaction.Rollback()
              $Mdb.Transaction = $null
            }
          }
          catch
          {
            $_.Exception.Message | Out-Host
          }
          break
        }
        ".nullvalue"
        {
          if ('"',"'",,'""',"''" -eq $ArgumentList[1])
          {
            $script:NullValue = $null
          }
          else
          {
            $script:NullValue =  $ArgumentList[1]
          }
        }
        ".width" {
          $script:Width = @()
          foreach ($w in $ArgumentList[1..($ArgumentList.Count - 1)])
          {
            [int]$Result = $null
            if ([int]::TryParse($w, [ref]$Result))
            {
              if ($Result -le 0)
              {
                break
              }
              $script:Width += $Result
            }
            else
            {
              break
            }
          }
          break
        }
        ".autosize"
        {
          $script:AutoSize = $OnOff
          break
        }
        ".echo" {
          $script:Echo = $OnOff
          break
        }
        ".headers"
        {
          $script:Headers = $OnOff
          break
        }
        ".wrap"
        {
          $script:Wrap = $OnOff
          break
        }
        ".show"
        {
          " autosize : $(if ($script:AutoSize) {"on"} else {"off"})" | Out-Host
          "     echo : $(if ($script:Echo) {"on"} else {"off"})"     | Out-Host
          "  headers : $(if ($script:Headers) {"on"} else {"off"})"  | Out-Host
          "nullvalue : `"$($script:NullValue)`""                     | Out-Host
          "    width : $Width"                                       | Out-Host
          "     wrap : $(if ($script:Wrap) {"on"} else {"off"})"     | Out-Host
        }
        ".help"
        {
          $script:HelpMessage | Out-Host
        }
        default
        {
          "error command : $DotCommand" | Out-Host
          continue
        }
      }

      #SQL Operation
      if ("go", "/", ";" -eq $Line -and -not $MultiLine)
      {
        continue
      }
      elseif ("go", "/", ";" -eq $Line)
      {
        try
        {
          if ($script:Echo)
          {
            $MultiLine | Out-Host
          }
          $Command = $Mdb | New-MdbCommand -CommandText $MultiLine
          $rows = $Command | Invoke-MdbCommand
          $rows | Out-MdbConsole | Out-Host
        }
        catch
        {
          if ($_.TargetObject.psobject.TypeNames -eq "MdbCommand.MdbError")
          {
            $MdbError = $_.TargetObject
            if ($MdbError.Message)
            {
              $MdbError.Message| Out-Host
            }
            else
            {
              $MdbError.ErrorRecord.Exception.InnerException.Message| Out-Host
            }
          }
          else
          {
            throw
          }
        }
        finally
        {
          $MultiLine = $null
          if ($Command)
          {
            $Command.Dispose()
          }
        }
      }
      elseif ($Line.Trim() -or (-not $Line.Trim() -and $MultiLine))
      {
         $MultiLine += "`n$Line"
      }
    }
    While (-not $ExitStatus)
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
function Out-MdbConsole
{
  [CmdletBinding()]
  param(
    [parameter(Mandatory=$true,
               ValueFromPipeline=$true)]
    [AllowNull()]
    [psobject]
    $InputObject
  )
  begin
  {
    $DoProcess = $false
    [object[]]$Property = @()
    if ($script:Width -or $script:NullValue)
    {
      function Quote
      {
        param([string]$s)
        "'$(if ($s.IndexOf("'")){$s.Replace("'", "''")} else {$s})'"
      }
      $Names = @(${script:MdbCommand.Schema} | % {$_.ColumnName})
      for ($i=0; $i -lt $Names.Count; $i++)
      {
        $Name = $Names[$i]
        if ($script:NullValue)
        {
          $Expr = [scriptblock]::Create(
            "if (`$_.$(Quote $Name) -is [dbnull])" + `
            "{$(Quote $script:NullValue)} " + `
            "else {`$_.$(Quote $Name)}"
          )
        }
        else
        {
          $Expr = "$Name"
        }
        if ($script:Width -and ($i -lt $script:Width.Count))
        {
          $Property += @{l=$Name; e=$Expr; width=$script:Width[$i]}
        }
        else
        {
          $Property += @{l=$Name; e=$Expr}
        }
      }
    }
    $ht = @{}
    if ($Property)
    {
      $ht.Add("Property", @($Property))                         | Out-Null
    }
    $ht.Add("AutoSize", [switch]$script:AutoSize)               | Out-Null
    $ht.Add("Wrap", [switch]$script:Wrap)                       | Out-Null
    $ht.Add("HideTableHeaders", [switch](-not $script:Headers)) | Out-Null
    
    $global:prop = $ht
    
    $WrappedCmd = $ExecutionContext.InvokeCommand.GetCmdlet("Format-Table")
    $ScriptCmd = { & $WrappedCmd @ht }
    $SteppablePipeline = $ScriptCmd.GetSteppablePipeline()
    $SteppablePipeline.Begin($pscmdlet)
    $DoProcess = $true
  }
  process
  {
    if ($DoProcess)
    {
       $SteppablePipeline.Process($_)
    }
  }
  end
  {
    $SteppablePipeline.End()
  }
}
