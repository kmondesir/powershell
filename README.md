# new-cabinet

Creating cabinet files to be utilized in provisioning packages

## Why 

To improve my understanding on writing PowerShell scripts as well as develop a style for that makes it easier to understand intention.

## Creating cabinet file within PowerShell

You can create a cabinet file within Powershell in a variety of ways. This script creates a ddf file which is essential in creating cabinet files as it describes the contents to what is included by the makecab command. This program is part of Windows and can be executed from the command line. 

e.g. `.\new-cab.ps1 -Source C:\bin\cool`

Will create a cabinet file in the script's current directory (default) that includes all items in the found in the cool directory. 

e.g. `.\new-cab.ps1 -Name Cool.cab -Source C:\bin\great`

Will create a cabinet file with the name Cool.cab script's current directory that contains the items from the great directory

e.g. `.\new-cab.ps1 -Name Cool.cab -Source C:\bin\great -Destination C:\bin\greater`

Will create an cab file with the name Cool.cab in the greater directory from the contents of the great directory

*Please note that you can change the name of the script file.*

## The original script can be found here 

[Virtual Engine] (https://virtualengine.co.uk/creating-cab-files-with-powershell/)
