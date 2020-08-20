Function Remove-ADComputerObject
{
	<#

	.NOTES
	Version: 1.0
	Author: Kino Mondesir
		
	.SYNOPSIS
	Remove Active Directory Object

	.DESCRIPTION
	Removes Active Directory Object after connecting with One Identity Service 

	.PARAMETER Name
	Mandatory. Name of the system you wish to remove

	.PARAMETER Service
	Optional. The Active Roles service that you must connect with, default value provided

	.PARAMETER Cred
	Mandatory. Credentials with permissions to access Active Roles service. This requires CORP\username-a credentials. 

	.EXAMPLE
	Remove-ADComputerObject -Name 'this server' -Cred Get-Credential

	.INPUTS
	3 inputs, 2 Mandatory and 1 Optional

	.OUTPUTS
	N/A

	.LINK
	https://www.youtube.com/watch?v=0DNXtRK187A
	Build and Throw Custom Exception Classes in PowerShell
	#>
	[CmdletBinding()]
	param 
	(
		[Parameter(HelpMessage="Hostname", Position = 0, Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='Name')]
		[ValidateNotNullOrEmpty()]
		[Alias('hostname')] 
		[string]$name,

		[Parameter(HelpMessage="Credentials passed in", Position = 1, Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
		[ValidateNotNullOrEmpty()]
		[pscredential]$cred,

		[Parameter(HelpMessage="Active Roles service to connect with", Position = 2, Mandatory=$false, ValueFromPipeline=$true)]
		[ValidateNotNullOrEmpty()]
		[string]$Service = $quest_server 
	)
	
	# classes

	class ObjectDoesNotExist : System.Exception
	{
		[string]$object
		ObjectDoesNotExist($object) : base("Object:$object does not exists!")
		{
			$this.object = $object
		}
	}

	Try
	{
		$system = Get-ADComputer -Filter 'Name -eq $Name'
		If ($system)
		{
			$system | Remove-ADComputer -Confirm:$false -Recursive
			return 0
		}
		else 
		{
			Throw [ObjectDoesNotExist]::New($name)
		}
	}
	catch [System.IndexOutOfRangeException]
	{
		$exception = $_.Exception
    Write-Error $exception.Message
		return -10
	}
	catch [ObjectDoesNotExist]
	{
		$exception = $_.Exception
    Write-Error $exception.Message
		return -20
	}
	catch 
	{
		$exception = $_.Exception
    Write-Error $exception.Message
		return -1
	}
}
