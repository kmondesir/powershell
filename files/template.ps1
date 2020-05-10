<#
.SYNOPSIS
  The purpose of the script. 

.DESCRIPTION
  Description of the script.

.PARAMETER Source
  Mandatory. Required attribute.

.PARAMETER Path
  Optional. Not required attribute.

  If not specified a default value will be applied

.EXAMPLE
  .\template.ps1 -Source "C:\Windows\Temp"

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

[CmdletBinding(DefaultParameterSetName = "Source")]
  Param( 
    [Parameter(HelpMessage = "Items to include in iso file.", Position = 0, Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = "Source")]
    [ValidateNotNullOrEmpty()]
    [string] $Source,

    [Parameter(HelpMessage = "Directory to put the file.", Position = 1, Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Source")]
    [ValidateScript({Test-Path -Path $_ -IsValid})]
    [string] $Path = $(Get-Location),
    
    [Parameter(HelpMessage = "Flag", Position = 2, Mandatory = $false, ValueFromPipelineByPropertyName = $true, ParameterSetName = "Flag")]
    [switch] $Flag 
  )

  Begin 
  {
    # Initialize variables
    # Check if shell is run as admin
    $TestRunAsAdmin = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")
    $Path = Get-Location 

    function Get-Timestamp 
    {
      return $(Get-Date).ToString("yyyy-MM-dd-HH:mm:ss.ffff")
    }

    function Message ([string]$statement)
    {
      return Get-Date -Format "Dyyyy-MM-ddTHH:mm:ss" + "::" + $statement
    }
  }

  Process
  {
    Start-Transcript -Path Get-Location
    # Execute
    try 
    {
      
    }
    catch 
    {
      # Potential errors
      Write-Error Message($_)
    }
    finally
    {
      Stop-Transcript
    }
  }

  End
  {
    # Clean up
    return 0
    Exit
  }