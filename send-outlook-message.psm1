Function new-message(){
<#
 
.SYNOPSIS
  Send a new email using the Outlook client
 
.DESCRIPTION
  The script uses the Microsoft Common Object Model, COM for short, for interacting with various Windows based applications. 
  This script in particular will be interacting with the Outlook client. 
 
.PARAMETER recipient
  This argument provides the recipient of the message. Multiple addresses can be included using ; as a separator
 
.PARAMETER carbonCopy
  This is an optional argument allowing you to add addition addresses to CC in a message
 
.PARAMETER subject
  This argument provides the subject for the message.
 
.EXAMPLE
  new-message creates a message with the default parameters excluding optional arguments
 
.NOTES
  Name: sendOutlookmessage
 
.LINK
  https://msdn.microsoft.com/en-us/VBA/Outlook-VBA/articles/olitemtype-enumeration-outlook
 
.LINK
  https://msdn.microsoft.com/en-us/vba/outlook-vba/articles/mailitem-object-outlook 
 
.LINK
  https://community.spiceworks.com/how_to/150253-send-mail-from-powershell-using-outlook
 
#>
 
[CmdletBinding()]
  Param(
    [Parameter(HelpMessage = "Recipient(s) added to the mail message. To include multiple email address please use ; as a separator.", Position = 0, Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
    [string[]] $recipient = 'morel.kevin@gmail.com',
 
    [Parameter(HelpMessage = "Additional addresses to be added to the mail message, use ; as a separator.", Position = 1, Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
    [string[]] $carbonCopy,
 
    [Parameter(HelpMessage = "If subject is not included a default string will include a timestamp", Position = 2, Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
    [string] $subject = ('Message sent at' + ' ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')),
 
    [Parameter(HelpMessage = "If body is not included a default string will include be provided instead.", Position = 3, Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
    [string] $body = 'Testing the default behavior of the script'
  )
  try {
    $Outlook = New-Object -ComObject Outlook.Application #Create an COM Object that exposes Outlook Application
    $Mail = $Outlook.CreateItem(0) #Reference the subset Mail 
 
    $Mail.To = $recipient
    $Mail.Subject = $subject
    $Mail.CC = $carbonCopy
    $Mail.Body = $body
 
    $Mail.Send()
  }
  catch {
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    Write-Output $ErrorMessage $FailedItem | Tee-Object -FilePath .\$date-errors.log
  }
  finally{
    $recipient = $null
    $subject = $null
    $carbonCopy = $null
    $body = $null
    $Outlook.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    Exit
  }
  
}
