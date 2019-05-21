#
# Windows PowerShell script for AD DS Deployment
#

New-Variable -Name 'domain' -Value 'dev.contoso.com' -Option Constant -Description 'domain controller'
New-Variable -Name 'netbios' -Value 'DEV' -Option Constant -Description 'netbios name'
New-Variable -Name 'computername' -Value 'dev-dc1-local' -Option Constant -Description 'hostname'
New-Variable -Name 'dns' -Value '127.0.0.1' -Option Constant -Description 'loopback address'
New-Variable -Name 'ipv4' -Value '192.168.1.14' -Option Constant -Description 'Change this address to the ip you want for the server'
New-Variable -Name 'defaultGateway' -Value '192.168.1.10' -Option Constant 'gateway address'
New-Variable -Name 'CIDR' -Value '24' -Option Constant -Description 'The network size or Prefixlength'

Install-windowsfeature -name AD-Domain-Services -IncludeManagementTools

Import-Module ADDSDeployment
Install-ADDSForest `
-CreateDnsDelegation:$false `
-DatabasePath "C:\Windows\NTDS" `
-DomainMode "WinThreshold" `
-DomainName $domain `
-DomainNetbiosName $netbios `
-ForestMode "WinThreshold" `
-InstallDns:$true `
-LogPath "C:\Windows\NTDS" `
-NoRebootOnCompletion:$false `
-SysvolPath "C:\Windows\SYSVOL" `
-Force:$true