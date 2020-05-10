<#
 
.SYNOPSIS
  Send a new email using the Outlook client
 
.DESCRIPTION
  The script uses the Microsoft Common Object Model, COM for short, for interacting with various Windows based applications. 
  This script in particular will be interacting with the Outlook client. 
 
.PARAMETER recipient
  Mandatory. Using an array of strings to provide multiple email addresses for the Outlook object
 
.PARAMETER carbonCopy
  Optional. Provides an array of emails addresses to be included in the cc line.
 
.PARAMETER subject
  Optional. Provides text to include in the subject.

.PARAMETER body
  Optional. Provides text to include in the body.
 
.EXAMPLE
  .\new-OutlookMessage -to "first.last@email.com"

.EXAMPLE
  .\new-OutlookMessage -to "first.last@email.com" -cc "john.smith@gmail.com"

.EXAMPLE
  .\new-OutlookMessage -to "first.last@email.com" -cc "john.smith@gmail.com" -bcc "jane.doe@gmail.com"

.EXAMPLE
  .\new-OutlookMessage -to "first.last@email.com" -cc "john.smith@gmail.com" -subject "Great Script" -body "Keep up the good work!"
 
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
    [Parameter(HelpMessage = "Recipient(s) added to the mail message. To include multiple email address please use ; as a separator.", Position = 0, Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
    [Alias("To")]
    [ValidateNotNullOrEmpty()]
    [string[]] $recipient,
 
    [Parameter(HelpMessage = "Additional addresses to be added to the mail message, use ; as a separator.", Position = 1, Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
    [Alias("Cc")]
    [string[]] $carbonCopy,

    [Parameter(HelpMessage = "Additional addresses to be added to the mail message, use ; as a separator.", Position = 1, Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
    [Alias("Bcc")]
    [string[]] $blindCopy,
 
    [Parameter(HelpMessage = "If subject is not included a default string will include a timestamp", Position = 2, Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $subject = ('Message sent at' + ' ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')),
 
    [Parameter(HelpMessage = "If body is not included a default string will include be provided instead.", Position = 3, Mandatory = $false, ValueFromPipelineByPropertyName = $true)]
    [Alias("Message")]
    [ValidateNotNullOrEmpty()]
    [string] $body = 'Testing the default behavior of the script'
  )
  Begin
  {
    $Outlook = New-Object -ComObject Outlook.Application #Create an COM Object that exposes Outlook Application
    $Mail = $Outlook.CreateItem(0) #Reference the subset Mail 
  }
  Process
  {
    Start-Transcript -Path $Path -Append
    try
    {
      $Mail.To = $recipient
      If ($carbonCopy){
        $Mail.Cc = $carbonCopy
      }
      elseif ($blindCopy) 
      {
        $Mail.Bcc = $blindCopy
      }
      $Mail.Subject = $subject
      $Mail.Body = $body
      $Mail.Send()
    }
    catch 
    {
      $ErrorMessage = $_.Exception.Message
      $FailedItem = $_.Exception.ItemName
      Write-Output $ErrorMessage $FailedItem | Tee-Object -FilePath .\$date-errors.log
    }
    finally 
    {
      $Outlook.Quit()
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
      Stop-Transcript
    }
  }
  End
  {
    return 0
    Exit
    # Completion Steps
  }