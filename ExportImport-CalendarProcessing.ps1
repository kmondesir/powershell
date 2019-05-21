<#
.SYNOPSIS
Export/Import Resource Booking Parameters.

.DESCRIPTION
Export/Import resource booking parameters for migrating resource mailbox
calendar processing information.  In many instances, calendar processing 
configuration is not persisted during mailbox migrations, so it is advisable
to export it prior to resource mailbox migration.

.PARAMETER Credential
Specify the credential used for the import or export action.

.PARAMETER Domains
Filter export results based on domain.

.PARAMETER Export
Enable export mode.

.PARAMETER ExportFile
Specify the file to export resource booking data to.

.PARAMETER ExportUri
Specify the endpoint for export.

.PARAMETER Identity
Specify an individual identity for import or export.

.PARAMETER Import
Enable import mode.

.PARAMETER ImportFile
Specify the file to use containing resource booking data.

.PARAMETER ImportUri
Specify the endpoint for import.

.PARAMETER UseExistingSession
Use the existing Exchange or Exchange Online session.

.EXAMPLE
.\ExportImport-CalendarProcessing.ps1 -Export -Credential (Get-Credential) -ExportFile CalendarProcessingData.csv

Connects to Office 365 using credential specified in (Get-Credential) and
exports calendar processing data for resource mailboxes to 
CalendarProcessingData.csv.

.EXAMPLE
.\ExportImport-CalendarProcessing.ps1 -Export -ExportUri https://onpremexchangeserver/powershell -UseExistingCredential -ExportFile CalendarProcessingData.csv

Connects to on-premises Exchange Server "onpremexchangeserver" using the
currently logged-in credential and exports calendar processing data for
resource mailboxes to CalendarProcessingData.csv.

.LINK
https://gallery.technet.microsoft.com/Export-and-Import-Calendar-123866af

.NOTES
Author: aaron.guilmette@microsoft.com

- 2019-03-11	Updated to remove Append parameter from Export-Csv.
- 2018-10-05	Updated to wrap resource delegates, requestoutofpolicy, 
				requestinpolicy, bookinpolicy attribute values in quotes.
- 2018-05-16	Updated Identity parameter to support an array input.
- 2018-04-05	Fixed typo in AllRequestOutOfPolicy.
				Added .Description to comment-based help.
				Added .Example to comment-based help.

#>
[CmdletBinding()] 
param ( 
# Export Parameters
[parameter(Mandatory=$false,ParameterSetName = "Export")] 
	[ValidateNotNullOrEmpty()] 
	[Switch]$Export,
	[String]$ExportUri = "https://outlook.office365.com/powershell-liveid/",
	[Array]$Domains,
[parameter(Mandatory=$true,ParameterSetName = "Export")]
	[string]$ExportFile,

# Import Parameters
[parameter(Mandatory=$false,ParameterSetName = "Import")] 
	[ValidateNotNullOrEmpty()] 
	[Switch]$Import,
	[String]$ImportUri = "https://outlook.office365.com/powershell-liveid/",
[parameter(Mandatory=$true,ParameterSetName = "Import")]
	[String]$ImportFile,
	
# Global Parameters
[parameter(Mandatory=$false,ParameterSetName = "Export")] 
[parameter(ParameterSetName = "Import")] 
	[System.Management.Automation.CredentialAttribute()]$Credential,
	[array]$Identity,
    [string]$LogFile = ".\CalendarExportImport.csv",
    [switch]$UseExistingSession
)

Function o365LogonExport
	{
	$ExportSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExportUri -Credential $Credential -Authentication Basic -AllowRedirection
	Import-PSSession $ExportSession
    }


Function o365LogonImport
	{
	$ImportSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ImportUri -Credential $Credential -Authentication Basic -AllowRedirection
	Import-PSSession $ImportSession
    }

Switch($PSCmdlet.ParameterSetName)
	{
	Export
	{
		# Connect to source Exchange environment
		If (!($UseExistingSession)) { o365LogonExport }
		
		# Create the domain filter
		If ($Domains)
		{
			$DomainsCount = $Domains.Count
			$i = 1
			$Filter = "{WindowsEmailAddress -like "
			Foreach ($Domain in $Domains)
			{
				If ($Domain.StartsWith("*"))
				{
					# Value already starts with an asterisk
				}
				Else
				{
					$Domain = "*" + $Domain
				}
				#$Filter = [scriptblock]::Create("{WindowsEmailAddress -like `"$Domain`"}")
				#$Filter = [scriptblock]::Create("`"$Domain`"")
				$Filter = $Filter + "`"$Domain`" -or WindowsEmailAddress -like "
			}
			$Filter = $Filter.Substring(0, $Filter.Length - 30)
			$Filter = $Filter + "}"
			Write-Host -NoNewline "Domain Filter is "; Write-Host -ForegroundColor Green $Filter
		}
		
		# Get the user list
		If ($Identity)
		{
			$Resources = @()
			foreach ($obj in $Identity) { [array]$Resources += (Get-Mailbox $obj) }
		}
		Else
		{
			Write-Host -ForegroundColor Green "Processing all mailboxes."
			#$cmd = "Get-Mailbox -ResultSize Unlimited -Filter { WindowsEmailAddress -like $Filter } -RecipientTypeDetails RoomMailbox,SharedMailbox,EquipmentMailbox"
			If ($Filter)
			{
				$cmd = "Get-Mailbox -ResultSize Unlimited -Filter $Filter -RecipientTypeDetails RoomMailbox,SharedMailbox,EquipmentMailbox"
			}
			Else
			{
				$cmd = "Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails RoomMailbox,SharedMailbox,EquipmentMailbox"
			}
			[array]$Resources = Invoke-Expression $cmd
		}
		
		Write-Host "Found $($Resources.Count) mailboxes."
		$Count = $Resources.Count
		
		$ResourceMailboxesSettings = @()
		$i = 1
		Foreach ($obj in $Resources)
		{
			Write-Host "Processing [$i / $Count] - $($obj.PrimarySmtpAddress)"
			$ResourceMailboxSetting = Get-CalendarProcessing $obj.PrimarySmtpAddress.ToString()
			Add-Member -InputObject $ResourceMailboxSetting -MemberType NoteProperty -Name "PrimarySmtpAddress" -Value $obj.PrimarySmtpAddress
			$ResourceMailboxesSettings += $ResourceMailboxSetting
			$i++
		}
		Write-Host "Making it pretty...please wait."
		$ResourceMailboxesSettings | Select PrimarySmtpAddress,`
											AutomateProcessing,`
											AllowConflicts,`
											BookingWindowInDays,`
											MaximumDurationInMinutes,`
											AllowRecurringMeetings,`
											EnforceSchedulingHorizon,`
											ScheduleOnlyDuringWorkHours,`
											ConflictPercentageAllowed,`
											MaximumConflictInstances,`
											ForwardRequestsToDelegates,`
											DeleteAttachments,`
											DeleteComments,`
											RemovePrivateProperty,`
											DeleteSubject,`
											AddOrganizerToSubject,`
											DeleteNonCalendarItems,`
											TentativePendingApproval,`
											EnableResponseDetails,`
											OrganizerInfo,`
											@{ n = 'ResourceDelegates'; e = { $val1 = @(); foreach ($obj in $_.ResourceDelegates) { $val1 += (Get-Recipient $obj).PrimarySmtpAddress }; $val1 = """" + ($val1 -join '";"') + """"; $val1 } },`
											@{ n = 'RequestOutOfPolicy'; e = { $val2 = @(); foreach ($obj in $_.RequestOutOfPolicy) { $val2 += (Get-Recipient $obj).PrimarySmtpAddress }; $val2 = """" + ($val2 -join '";"') + """"; $val2 } },`
											@{ n = 'BookInPolicy'; e = { $val3 = @(); foreach ($obj in $_.BookInPolicy) { $val3 += (Get-Recipient $obj).PrimarySmtpAddress }; $val3 = """" + ($val3 -join '";"') + """"; $val3 } },`
											@{ n = 'RequestInPolicy'; e = { $val4 = @(); foreach ($obj in $_.RequestInPolicy) { $val4 += (Get-Recipient $obj).PrimarySmtpAddress }; $val4 = """" + ($val4 -join '";"') + """"; $val4 } },`
		
											#@{n='ResourceDelegates';e={$val1=@();foreach ($obj in $_.ResourceDelegates){ $val1 += (Get-Recipient $obj).PrimarySmtpAddress} $val1 -join ";"}},`
											#@{n="RequestOutOfPolicy";e={$val2=@();foreach ($obj in $_.RequestOutOfPolicy){ $val2 += (Get-Recipient $obj).PrimarySmtpAddress} $val2 -join ";"}},`
											#@{n="RequestInPolicy";e={$val4=@();foreach ($obj in $_.RequestInPolicy){ $val4 += (Get-Recipient $obj).PrimarySmtpAddress} $val4 -join ";"}},`
											#@{n="BookInPolicy";e={$val3=@();foreach ($obj in $_.BookInPolicy){ $val3 += (Get-Recipient $obj).PrimarySmtpAddress} $val3 -join ";"}},`
											AllRequestOutOfPolicy,`
											AllBookInPolicy,`
											AllRequestInPolicy,`
											AddAdditionalResponse,`
											AdditionalResponse,`
											RemoveOldMeetingMessages,`
											AddNewRequestsTentatively,`
											ProcessExternalMeetingMessages,`
											RemoveForwardedMeetingNotifications `
		| Export-Csv -NoTypeInformation $ExportFile # -Append
		Write-Host -NoNewline "Export file is "; Write-Host -NoNewLine -ForegroundColor Green $ExportFile; Write-Host "."
		Get-PSSession | ? { $_.ConfigurationName -eq "Microsoft.Exchange" } | Remove-PSSession
		} # End Export
	Import
		{
		# Connect to target Exchange environment
        If (!($UseExistingSession)) { o365LogonImport }
        
        If ($Identity)
			{
			$Temp = Import-Csv $ImportFile
			[array]$ResourceMailboxSettings = $Temp | ? { $_.PrimarySmtpAddress -match $Identity }
			}
		Else
			{
			[array]$ResourceMailboxSettings = Import-Csv $ImportFile
			}
		$i = 1
        $Count = $ResourceMailboxSettings.Count
        Write-Host "Processing $($Count) mailboxes."
		Foreach ($Mailbox in $ResourceMailboxSettings)
			{
			Write-Host "Processing [$i / $Count] - $($Mailbox.PrimarySmtpAddress)"
            $RecipientType = (Get-Mailbox $Mailbox.PrimarySmtpAddress).RecipientTypeDetails
            $cmd = "Set-CalendarProcessing -Identity $($Mailbox.PrimarySmtpAddress)"
			If ($Mailbox.AutomateProcesing) { $AutomateProcessing = $Mailbox.AutomateProcessing; $cmd = $cmd + " -AutomateProcessing $AutomateProcessing" }
			If ($Mailbox.AllowConflicts) { $AllowConflicts = "`$"+$Mailbox.AllowConflicts; $cmd = $cmd + " -AllowConflicts $AllowConflicts" }
			If ($Mailbox.BookingWindowInDays) { $BookingWindowInDays = $Mailbox.BookingWindowInDays; $cmd = $cmd + " -BookingWindowInDays $BookingWindowInDays" }
			If ($Mailbox.MaximumDurationInMinutes) { $MaximumDurationInMinutes = $Mailbox.MaximumDurationInMinutes; $cmd = $cmd + " -MaximumDurationInMinutes $MaximumDurationInMinutes" }
			If ($Mailbox.AllowRecurringMeetings) { $AllowRecurringMeetings = "`$"+$Mailbox.AllowRecurringMeetings; $cmd = $cmd + " -AllowRecurringMeetings $AllowRecurringMeetings" }
			If ($Mailbox.EnforceSchedulingHorizon) { $EnforceSchedulingHorizon = "`$"+$Mailbox.EnforceSchedulingHorizon; $cmd = $cmd + " -EnforceSchedulingHorizon $EnforceSchedulingHorizon" }
			If ($Mailbox.ScheduleOnlyDuringWorkingHours) { $ScheduleOnlyDuringWorkingHours = "`$"+$Mailbox.ScheduleOnlyDuringWorkingHours; $cmd = $cmd + " -ScheduleOnlyDuringWorkingHours $ScheduleOnlyDuringWorkingHours" }
			If ($Mailbox.ConflictPercentageAllowed) { $ConflictPercentageAllowed = $Mailbox.ConflictPercentageAllowed; $cmd = $cmd + " -ConflictPercentageAllowed $ConflictPercentageAllowed" }
			If ($Mailbox.MaximumConflictInstances) { $MaximumConflictInstances = $Mailbox.MaximumConflictInstances; $cmd = $cmd + " -MaximumConflictInstances $MaximumConflictInstances" }
			If ($Mailbox.ForwardRequestsToDelegate) { $ForwardRequestsToDelegate = "`$"+$Mailbox.ForwardRequestsToDelegate; $cmd = $cmd + " -ForwardRequestsToDelegate $ForwardRequestsToDelegate" }
			If ($Mailbox.DeleteAttachments) { $DeleteAttachments = "`$"+$Mailbox.DeleteAttachments; $cmd = $cmd + " -DeleteAttachments $DeleteAttachments" }
			If ($Mailbox.DeleteComments) { $DeleteComments = "`$"+$Mailbox.DeleteComments; $cmd = $cmd + " -DeleteComments $DeleteComments" }
			If ($Mailbox.RemovePrivateProperty) { $RemovePrivateProperty = "`$"+$Mailbox.RemovePrivateProperty; $cmd = $cmd + " -RemovePrivateProperty $RemovePrivateProperty" }
			If ($Mailbox.DeleteSubject) { $DeleteSubject = "`$"+$Mailbox.DeleteSubject; $cmd = $cmd + " -DeleteSubject $DeleteSubject" }
			If ($Mailbox.AddOrganizerToSubject) { $AddOrganizerToSubject = "`$"+$Mailbox.AddOrganizerToSubject; $cmd = $cmd + " -AddOrganizerToSubject $AddOrganizerToSubject" }
			If ($Mailbox.DeleteNonCalendarItems) { $DeleteNonCalendarItems = "`$"+$Mailbox.DeleteNonCalendarItems; $cmd = $cmd + " -DeleteNonCalendarItems $DeleteNonCalendarItems" }
			If ($Mailbox.TentativePendingApproval) { $TentativePendingApproval = "`$"+$Mailbox.TentativePendingApproval; $cmd = $cmd + " -TentativePendingApproval $TentativePendingApproval" }
			If ($Mailbox.EnableResponseDetails) { $EnableResponseDetails = "`$"+$Mailbox.EnableResponseDetails; $cmd = $cmd + " -EnableResponseDetails $EnableResponseDetails" }
			If ($Mailbox.OrganizerInfo) { $OrganizerInfo = "`$"+$Mailbox.OrganizerInfo; $cmd = $cmd + " -OrganizerInfo $OrganizerInfo" }
			If ($Mailbox.RequestOutOfPolicy -and $Mailbox.RequestOutOfPolicy -notmatch '^\"\"$') { $RequestOutOfPolicy = $Mailbox.RequestOutOfPolicy.Replace(";",","); $cmd = $cmd + " -RequestOutOfPolicy $RequestOutOfPolicy" }
			If ($Mailbox.AllRequestOutOfPolicy) { $AllRequestOutOfPolicy = "`$"+$Mailbox.AllRequestOutOfPolicy; $cmd = $cmd + " -AllRequestOutOfPolicy $AllRequestOutOfPolicy" }
			If ($Mailbox.BookInPolicy -and $Mailbox.BookInPolicy -notmatch '^\"\"$') { $BookInPolicy = $Mailbox.BookInPolicy.Replace(";",",").Replace(",,",","); $cmd = $cmd + " -BookInPolicy $BookInPolicy" }
			If ($Mailbox.AllBookInPolicy) { $AllBookInPolicy = "`$"+$Mailbox.AllBookInPolicy; $cmd = $cmd + " -AllBookInPolicy $AllBookInPolicy" }
			If ($Mailbox.RequestInPolicy -and $Mailbox.RequestInPolicy -notmatch '^\"\"$') { $RequestInPolicy = $Mailbox.RequestInPolicy.Replace(";",","); $cmd = $cmd + " -RequestInPolicy $RequestInPolicy" }
			If ($Mailbox.AllRequestInPolicy) { $AllRequestInPolicy = "`$"+$Mailbox.AllRequestInPolicy; $cmd = $cmd + " -AllRequestInPolicy $AllRequestInPolicy" }
			If ($Mailbox.AddAdditionalResponse) { $AddAdditionalResponse = "`$"+$Mailbox.AddAdditionalResponse; $cmd = $cmd + " -AddAdditionalResponse $AddAdditionalResponse" }
			If ($Mailbox.AdditionalResponse) 
				{ 
				#$AdditionalResponse = $Mailbox.AdditionalResponse.Replace("\","\\").Replace(".","\.").Replace("^","\^").Replace("$","\$").Replace("*","\*").Replace("+","\+").Replace("-","\-").Replace("?","\?").Replace("(","\(").Replace(")","\)").Replace("[","\[").Replace("]","\]").Replace("{","\{").Replace("}","\}").Replace("|","\|").Replace("<","\<").Replace(">","\>").Replace(":","\:").Replace("@","\@").Replace("/","\/").Replace("'","\'")
				$AdditionalResponse = $Mailbox.AdditionalResponse
				$cmd = $cmd + " -AdditionalResponse ""$AdditionalResponse"" "
				}
			If ($Mailbox.RemoveOldMeetingMessages) { $RemoveOldMeetingMessages = "`$"+$Mailbox.RemoveOldMeetingMessages; $cmd = $cmd + " -RemoveOldMeetingMessages $RemoveOldMeetingMessages" }
            If ($Mailbox.RemoveForwardedMeetingNotifications) { $RemoveForwardedMeetingNotifications = "`$"+$Mailbox.RemoveForwardedMeetingNotifications; $cmd = $cmd + " -RemoveForwardedMeetingNotifications $RemoveForwardedMeetingNotifications" }
			# Params that fail if mailbox is not configured as "RoomMailbox" or "EquipmentMailbox"
            If ($Mailbox.AddNewRequestsTentatively) 
                {
                If ($RecipientType -match "Equipment|Room")
                    { $AddNewRequestsTentatively = "`$"+$Mailbox.AddNewRequestsTentatively; $cmd = $cmd + " -AddNewRequestsTentatively $AddNewRequestsTentatively" }
                Else
                    {
                    Write-Host -ForegroundColor Red "Object $($Mailbox.PrimarySmtpAddress) has value for AddNewRequestsTentatively, but is not configured as a resource mailbox."
                    $data = "Object $($Mailbox.PrimarySmtpAddress) has value for AddNewRequestsTentatively, but is not configured as a resource mailbox."
                    $data | Out-File $LogFile -Append
                    }
                }
            If ($Mailbox.ProcessExternalMeetingMessages) 
                {
                If ($RecipientType -match "Equipment|Room")
                    { 
                    $ProcessExternalMeetingMessages = "`$"+$Mailbox.ProcessExternalMeetingMessages; $cmd = $cmd + " -ProcessExternalMeetingMessages $ProcessExternalMeetingMessages" 
                    }
                Else
                    {
                    Write-Host -ForegroundColor Red "Object $($Mailbox.PrimarySmtpAddress) has value for ProcessExternalMeetingMessages, but is not configured as a resource mailbox."
                    $data = "Object $($Mailbox.PrimarySmtpAddress) has value for ProcessExternalMeetingMessages, but is not configured as a resource mailbox."
                    $data | Out-File $Logfile -Append
                    }
                }
            If ($Mailbox.ResourceDelegates -and $Mailbox.ResourceDelegates -notmatch '^\"\"$')
                { 
                If ($RecipientType -match "Equipment|Room")
                    { 
                    $ResourceDelegates = $Mailbox.ResourceDelegates.Replace(";",","); $cmd = $cmd + " -ResourceDelegates $ResourceDelegates" 
                    }
                Else
                    {
                    Write-Host -ForegroundColor Red "Object $($Mailbox.PrimarySmtpAddress) has value for ResourceDelegates, but is not configured as a resource mailbox."
                    $data = "Object $($Mailbox.PrimarySmtpAddress) has value for ResourceDelegates, but is not configured as a resource mailbox."
                    $data | Out-File $Logfile -Append
                    }
                }			
			#Write-Host "The command to be executed is: $cmd"
			Invoke-Expression $cmd
            $i++
			}
		If (!($UseExistingSession)) { Get-PSSession | ? { $_.ConfigurationName -eq "Microsoft.Exchange" } | Remove-PSSession }
        Write-Host "Finished importing!"
		} # End Import
	} # End Switch