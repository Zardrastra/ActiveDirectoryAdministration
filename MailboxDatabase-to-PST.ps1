#Create table of NT Account Aliases which match the rule "AccountDisabled and which are "UserMailbox"
$Mailboxes = Get-Mailbox -ResultSize Unlimited | where-object { $_.ExchangeUserAccountControl -like "AccountDisabled" -and $_.RecipientTypeDetails -eq 'UserMailbox' } | select-object alias
$Global:RunningStatuses = "Completed",  "InProgress", "Failed", "Queued"
#Variables controling backup folder directory
# $ExcDomainControler is the Domain Controller
# $LocalShare is the Exchange accessable folder where each backup is to be written to.
# $RemoteShare is the location of the offsite share where the script is to upload the fully completed PST backup file to.

[string] $Global:ExcDomainControler = (get-addomain).pdcemulator
[string] $Global:RemoteShare = #Path to external storage share (i.e. Azure Cloud -example "\\something.contoso.com\Archive\PST\"
[string] $Global:LocalShare = #Path to local network share - Example: "\\Server\ShareName\Output\"

#Function used to print status of the export attempt and to delay copying the mailbox until the process has completed, this will also skip mailboxes which are in "completed" status.
function ExchangeWaitLoop {
$global:starttime = get-date
clear-host

    
    [string] $Global:filepath = $Global:LocalShare + $pstname + ".pst"
    $Global:InitialPST = Get-MailboxExportRequest -Mailbox $pstname | Get-MailboxExportRequestStatistics
    $Global:ExportStatus = $InitialPST.Status
    $RunningStatuses = "Completed", "InProgress", "Failed", "Queued"
    $NTAccount = Get-ADuser pstname
    $Global:EmployeeName = $NTAccount.Name

    if ($RunningStatuses -notcontains $ExportStatus){
@"
Beginning PST Backup process for $EmployeeName

******************************************
"@
                New-MailboxExportRequest -DomainController $ExcDomainControler -mailbox $pstname -FilePath $filepath
    }
    [Bool]$global:ProcessActive = $True


    while ($ProcessActive -eq $True){
        $global:ProcessingQueue = Get-MailboxExportRequest -Mailbox $pstname | Get-MailboxExportRequestStatistics 

        #Find mailbox status after running backup request 
        $Global:PercentComplete = $ProcessingQueue.PercentComplete
        $global:ExportStatus = $ProcessingQueue.Status

		    if ($ExportStatus -eq "Completed"){
            #Test to check if the mailbox process is complete, ends loop by setting Process Active flag to false
		    clear-host
		    [Bool]$global:ProcessActive = $False
			[Bool]$PSTExists = test-path $filepath
			$ProcessingQueue = Get-MailboxExportRequest -Mailbox $pstname | Get-MailboxExportRequestStatistics 

			#Find mailbox status after running backup request 
			$PercentComplete = $ProcessingQueue.PercentComplete
			$ExportStatus = $ProcessingQueue.Status
				if($ExportStatus -eq "Completed"){
					if ($PSTExists -eq $True){
						"Mailbox copy for User $pstname completed, copying to remote share."
                        Start-Sleep -s 15
						move-item $filepath $RemoteShare
						#clear-host
						}
				}
				Elseif($PSTExists -eq $False){
				"No PST for user $pstname exists, skipping to next user"
				}
		    }
            Elseif($ExportStatus -eq "Failed"){

            $PSTCopyErr = "PST Process For user $pstname failed"
            write-host $PSTCopyErr
            $loglocation = $LocalShare + "output.txt"
            $PSTCopyErr >> $loglocation
            "Press any key to continue"
            pause
            }

		    else {
#Else function to generate status message of the ongoing backup with timestamps
		    $currentTime = get-date
            Start-Sleep -s 5
		    clear-host
            $CurrentTimetext = "Current time:" + " " + $currentTime.ToShortTimeString()
@"
Awaiting mailbox backup for:

$EmployeeName

Start time: $starttime


Current time: $CurrentTimetext
Percent completed: $PercentComplete %
"Refreshing backup progress in 2 minutes"
"@
Write-Host `a
Start-Sleep -s 120
clear-host
"Refreshing backup progress..."
		}
	}
    "Next mailbox backup in 5 seconds"
    Start-Sleep -s 5
    clear-host
}

foreach ($mailbox in $Mailboxes) {
$global:pstname = $mailbox.alias
ExchangeWaitLoop
}
