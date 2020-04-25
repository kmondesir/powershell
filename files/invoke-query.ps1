<#
.SYNOPSIS
  Execute SQL Query. 

.DESCRIPTION
  Sends SQL Commands to a database.

.PARAMETER Provider
  Optional. The data provider to use. If no provider is specified then Access is used. 

.PARAMETER Source
  Mandatory. Path of the database.

.PARAMETER Database
  Mandatory. Name of the database.

.PARAMETER Command
  Mandatory. SQL Query.

.PARAMETER Timeout
  Optional. Timeout in seconds to wait for the query to complete. Default is 60 seconds.

.PARAMETER Credential
  Optional. Credentials to use in connection if any.

.EXAMPLE
  .\invoke-query.ps1 -Source "C:\Datasource" -Command "SELECT * FROM Table"

.EXAMPLE
  .\invoke-query.ps1 -Source "C:\Datasource" -Database "Northwind" -Command "SELECT * FROM Table"

.EXAMPLE
  .\invoke-query.ps1 -Source "C:\Datasource" -Database "Northwind" -Command "SELECT * FROM Table" -Timeout 120

.EXAMPLE
  .\invoke-query.ps1 -Provider "MSOLEDBSQL" -Source "C:\Datasource" -Database "Northwind" -Command "SELECT * FROM Table" -Timeout 120

.EXAMPLE
  .\invoke-query.ps1 -Provider "MSOLEDBSQL" -Source "C:\Datasource" -Database "Northwind" -Command "SELECT * FROM Table" -Timeout 120 -Authentication "Integrated Security=SSPI;"

.INPUTS
  None. 

.OUTPUTS
  The file that is written to disk

.LINK
  https://gallery.technet.microsoft.com/scriptcenter/New-ISOFile-function-a8deeffd 

.NOTES
  Addition information about the script
  Version: 1.0
  Author: Myself
#>

[CmdletBinding()]
  Param( 
    [Parameter(HelpMessage = "Database provider", Position = 0, Mandatory = $false, ValueFromPipeline = $true)]
    [string] $Provider = 'Microsoft.ACE.OLEDB.12.0',

    [Parameter(HelpMessage = "Location of database", Position = 1, Mandatory = $true, ValueFromPipeline = $true)]
    [string] $Source,

    [Parameter(HelpMessage = "Database name", Position = 2, Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
    [string] $Name,

    [Parameter(HelpMessage = "Sequel statement", Position = 3, Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
    [string[]] $Queries,

    [Parameter(HelpMessage = "Timeout value", Position = 4, Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
    [int] $Timeout = 60,

    [Parameter(HelpMessage = "Authentication method", Position = 5, Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
    [string] $Authentication = "Integrated Security=SSPI;",

    [Parameter(HelpMessage = "User credentials", Position = 6, Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
    [securestring]$Cred
  )

  Begin 
  {
    # Initialize variables
    Set-StrictMode -Version 3
    # connection string
    [string] $connectionString = "Provider=$provider;Data Source=$Source"
    Write-Verbose $connectionString
  }

  Process
  {
    # Execute commands
    try 
    {
      # create connection object
      $connectionObject = New-Object System.Data.OleDb.OleDbConnection $connectionString
      $connectionObject.Open

      foreach($query in $queries)
      {
        # create query object
        $command = New-Object Data.OleDb.OleDbCommand $query, $connectionObject
        $command.CommandTimeout = $timeout

        # create adapter object based on query
        $adapter = New-Object System.Data.OleDb.OleDbDataAdapter $command

        # create empty dataset object
        $dataset = New-Object System.Data.Dataset

        # fill empty dataset object
        [void] $adapter.Fill($dataSet)

        # return all rows from memory
        return $dataset
      }
    }
    catch [System.SystemException.OleDbConnection]
    {
      Write-Error "Unable to open connection"
    }
    catch [System.SystemException.OleDbCommand]
    {
      Write-Error "Unable to open command"
    }
    catch
    {
      Write-Error "Unknown Error"
    }
    finally
    {
      # Close connections
      $connectionObject.Close()
    }
  }

  End
  {
    # Exit Script
    Exit
  }