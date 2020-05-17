<#

.SYNOPSIS
  Creates a hyper-v machine

.DESCRIPTION
  Creates a virtual machine with predefined attributes

.PARAMETER Title
  Mandatory. The name of the virtual machine. 

.PARAMETER Memory
  Optional. The amount of memory in megabytes. If no amount given 2GB will be used. 

.PARAMETER Path
  Optional. Location of the virtual machine. If left blank it will use current profile documents folder.

.PARAMETER Disk
  Optional. Location of the virtual disk. If left blank it will use the documents folders and append the title name.

.PARAMETER Size
  Optional. Size of the disk in gigabytes. If no amount given 20GB will be used.

.PARAMETER Switch
  Optional. The virtual switch that the vm is attached to.

.PARAMETER ISO
  Optional. The ISO file the system boots from. 

.EXAMPLE
  .\new-hyper-v.ps1 -Title "PowerShell"

.EXAMPLE
  .\new-hyper-v.ps1 -Title "PowerShell" -Memory 4GB

.EXAMPLE
  .\new-hyper-v.ps1 -Title "PowerShell" -Path "C:\Users\Admin\Documents\HyperV"

.EXAMPLE
  .\new-hyper-v.ps1 -Title "PowerShell" -Disk "C:\Users\Admin\Documents\HyperV\$title"

.EXAMPLE
  .\new-hyper-v.ps1 -Title "PowerShell" -Size 20GB

.EXAMPLE
  .\new-hyper-v.ps1 -Title "PowerShell" -Switch "Bridged Virtual Switch"

.EXAMPLE
  .\new-hyper-v.ps1 -Title "PowerShell" -ISO "\\images\win10.iso"

.INPUTS
  None. 

.OUTPUTS
  The file that is written to disk.

.LINK
  https://www.youtube.com/watch?v=HQF5Gr8tlks

.LINK
  https://docs.microsoft.com/en-us/powershell/module/hyper-v/new-vm?view=win10-ps 

.NOTES
  Addition information about the script
  Version: 1.0
  Author: Kino Mondesir
#>

[CmdletBinding()]
  Param
  ( 
    [Parameter(HelpMsg = "Name of the virtual machine", Position = 0, Mandatory = $false, ValueFromPipelineByPropertyName = $false)]
    [Alias("Name")]
    [string] $Title = $(Get-Date).ToString("yyyyMMdd-HHmmss.ffff"),

    [Parameter(HelpMsg = "Memory Startup Bytes", Position = 1, Mandatory = $false, ValueFromPipelineByPropertyName = $false)]
    [Alias("RAM")]
    [string] $Memory = 2048MB,

    [Parameter(HelpMsg = "Virtual Machine Location", Position = 2, Mandatory = $false, ValueFromPipelineByPropertyName = $false)] 
    [string] $Path = (Join-Path -Path $env:USERPROFILE -Childpath "documents\hyper-v\$Title"),

    [Parameter(HelpMsg = "Virtual Disk Location", Position = 3, Mandatory = $false, ValueFromPipelineByPropertyName = $false)] 
    [string] $Disk = (Join-Path -Path $env:USERPROFILE -Childpath "documents\hyper-v\$Title\$Title.vhdx"),

    [Parameter(HelpMsg = "Virtual Disk Size", Position = 4, Mandatory = $false, ValueFromPipelineByPropertyName = $false)] 
    [Alias("Storage")]
    [string] $Size = 20GB,

    [Parameter(HelpMsg = "Network Switch", Position = 5, Mandatory = $false, ValueFromPipelineByPropertyName = $false)]
    [string] $Switch = "Bridged Virtual Switch",

    [Parameter(HelpMsg = "ISO File", Position = 6, Mandatory = $false, ValueFromPipelineByPropertyName = $false)]
    [Alias("Image")]
    [string] $ISO,

    [Parameter(HelpMsg = "General information about the virtual machine", Position = 7, Mandatory = $false, ValueFromPipelineByPropertyName = $false)] 
    [Alias("Memo")]
    [string] $Note
  )

  Begin 
  {
    $VerbosePreference = 'Continue'
    # Check if shell is run as admin
    $TestRunAsAdmin = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")
    $vmExists = [bool](get-vm -name $title -ErrorAction SilentlyContinue)
    $isHyperVEnabled = [bool](Get-WindowsOptionalFeature -FeatureName Microsoft-Hyper-V-All -Online)
    $Path = Convert-Path -Path . 
    function Get-Timestamp 
    {
      return $(Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    }

    function Msg ([string]$statement)
    {
      $now = Get-Date -Format 'yyyy-mm-dd HH:mm:ss'
      return $now + "::" + $statement
    }
  }

  Process
  {
    Start-Transcript -Path $Path -Append
    Try
    {
      If ($TestRunAsAdmin)
      {
        If (!$isHyperVEnabled)
        {
          # Checks if Hyper-V manager is enabled
          Write-Verbose -Message (Msg "Hyper-V manager is NOT enabled.")
        }
        elseif ($vmExists) 
        {
          # Checks if VM is already created
          Write-Verbose -Message (Msg "VM already exists. Please use another name.")
        }
        Else
        {
          # Create virtual machine
          Write-Verbose -Message (Msg "Create virtual machine:$Title with $Memory of memory and $Size of storage")
          New-VM -Name $Title -MemoryStartupBytes $Memory -Path $Path -NewVHDPath $Disk -NewVHDSizeBytes $Size 
          If ($ISO) 
          {
            # Add ISO image to boot from
            Add-VMDvdDrive -VMName $Title -Path $ISO
            Write-Verbose -Message (Msg"Added ISO:$ISO")
          }
          Elseif ($Note) 
          {
            # Add Notes to VM
            Set-VM -Name $Title -Notes $Note
            Write-Verbose -Message (Msg"Added Note:$Note")
          }
        } 
      }
      else
      {
        # Hyper V Manager is not enabled
        # Path is not accessible
        $exception = $_.Exception.Message
        Write-Host (Msg "Please start script with Administrative Priviledges: $exception")
        return $_.Exception.ID
        Exit
      }
    }
    Catch 
    {
      $exception = $_.Exception.Message
      Write-Error (Msg "Unknown: $exception")
      return $_.Exception.ID
      Exit
    }
    Finally
    {
      Stop-Transcript
    }
  }

  End
  {
    return 0
    Exit
  }