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
    [Parameter(HelpMessage = "Name of the virtual machine", Position = 0, Mandatory = $false, ValueFromPipeline = $false)]
    [string] $Title = $(Get-Date).ToString("yyyyMMdd-HHmmss.ffff"),

    [Parameter(HelpMessage = "Memory Startup Bytes", Position = 1, Mandatory = $false, ValueFromPipeline = $false)]
    [Alias("RAM")]
    [string] $Memory = 2GB,

    [Parameter(HelpMessage = "Virtual Machine Location", Position = 2, Mandatory = $false, ValueFromPipeline = $false)] 
    [string] $Path = (Join-Path -Path $env:USERPROFILE -Childpath "documents\hyper-v"),

    [Parameter(HelpMessage = "Virtual Disk Location", Position = 3, Mandatory = $false, ValueFromPipeline = $false)] 
    [string] $Disk = (Join-Path -Path $env:USERPROFILE -Childpath ("documents\hyper-v\$Title.vhdx")),

    [Parameter(HelpMessage = "Virtual Disk Size", Position = 4, Mandatory = $false, ValueFromPipeline = $false)] 
    [Alias("Storage")]
    [string] $Size = 20GB,

    [Parameter(HelpMessage = "Network Switch", Position = 5, Mandatory = $false, ValueFromPipeline = $false)]
    [string] $Switch = "Bridged Virtual Switch",

    [Parameter(HelpMessage = "ISO File", Position = 6, Mandatory = $false, ValueFromPipeline = $false)]
    [Alias("Image")]
    [string] $ISO,

    [Parameter(HelpMessage = "General information about the virtual machine", Position = 7, Mandatory = $false, ValueFromPipeline = $false)] 
    [Alias("Memo")]
    [string] $Note = "$Title created with `n memory:$memory `n storage:$size"
  )

  Begin 
  {
    # Check if shell is run as admin
    $TestRunAsAdmin = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")
  }

  Process
  {
    Try
    {
      If ($TestRunAsAdmin)
      {
        Write-Verbose "Create $Title virtual machine with $Memory of memory and $Size of storage"
        # Create virtual machine
        New-VM -Name $Title -MemoryStartupBytes $Memory -Path $Path -NewVHDPath $Disk -NewVHDSizeBytes $Size -Notes $Note
                
        Write-Verbose "Start $Title virtual machine"
        If ($ISO)
        {
          # Add ISO image to boot from
          Write-Verbose "Add $ISO file"
          Add-VMDvdDrive -VMName $Title -Path $ISO
        }
      }
      else
      {
        # Hyper V Manager is not enabled
        # Path is not accessible
        Write-Error "Please run as shell as admin"
      }
    }
    Catch
    {
      Write-Error "Unknown error"
    }
  }

  End
  {
    # Clean up task
    Exit
  }