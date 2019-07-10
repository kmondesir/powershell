<#
2019-6-3 KMM: First version

Author: Kino Mondesir
Purpose: Locate users whose accounts have expired but are NOT Permanently Disabled.

#>

Set-Variable -name 'recipient' -option Constant -scope global -value 'itsupport@nbrem.no'
Set-Variable -name 'send' -option Constant -scope global -value 'noreply@nbrem.no'
Set-Variable -name 'SMTP' -option Constant -scope global -value 'relay.nbrem.no'
$ErrorActionPreference = Stop
$date = Get-Date
$primaryGroup = "_PermanentlyDisabledUsers"

Try
{
    $group = Get-ADGroup -Identity $primaryGroup -properties @("primaryGroupToken")
    $expired = Search-ADAccount -AccountExpired -UsersOnly 
    $permanentlyDisabled = Get-ADGroupMember -Identity $primaryGroup -Recursive
    $results = $expired | Where-Object { $_.SID -Notin $permanentlyDisabled.SID } # Select all users whose account has expired but NOT in the Permanently Disabled group
    If ($null -ne $results)
    {
        $results | ForEach-Object { Add-ADGroupMember -Identity $primaryGroup -Members $_.DistinguishedName -Confirm:$false } # Add users to Disabled Group
        $results | ForEach-Object { Set-ADUser -Identity $_ -Replace @{primarygroupid = $group.primaryGroupToken}} # Change primary group to Disabled Group
        $results | ForEach-Object { Get-ADUser -Identity $_ -Properties MemberOf } | ForEach-Object {
        $_.MemberOf | Remove-ADGroupMember -Members $_.DistinguishedName -Confirm:$false } # Remove all groups from users
        $results | ForEach-Object { Set-ADUser -Identity $_.DistinguishedName -Enabled $false -Confirm:$true } # Disable user accounts
        $results | Export-Clixml -Path .\$date-results.xml
        $results | Export-CSV -Path .\$date-results.csv
        send-mailmessage -to $recipient -from $send -subject 'Offboarding script success' -smtpServer $SMTP -Attachments .\$date-results.csv
    }
    Else
    {
        Exit
    }
}
Catch
{
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    Write-Host $ErrorMessage $FailedItem | Out-File .\$date-errors.log
    send-mailmessage -to $recipient -from $send -subject 'Offboarding script failure' -smtpServer $SMTP -Attachments .\$date-errors.log
    # Would like to capture errors and send an email to Fresh Service to generate a ticket
}
Finally
{
    $expired = $null
    $permanentlyDisabled = $null
    $results = $null
    $date = $null
}