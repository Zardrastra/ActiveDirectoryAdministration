#Use case: This script allows the backup of a disabled mailbox from a mailbox database to a share that is writable
#Once downloaded this script will then copy this to another storage location
#this is useful in instances where the mail servers do not have direct write access to the final storage location
#and are writing to a space limited location
#(servers behind a DMZ for example)

#Collect mailbox aliases which match the rule "AccountDisabled and which are "UserMailbox"
Write-host "Collecting Disabled Mailbox information for backup"
$Mailboxes = Get-Mailbox -ResultSize Unlimited | where-object { $_.ExchangeUserAccountControl -like "AccountDisabled" -and $_.RecipientTypeDetails -eq 'UserMailbox' } | select-object alias
clear-host
$Global:RunningStatuses = "Completed",  "InProgress", "Failed", "Queued"
#Variables controling backup folder directory
# $ExcDomainControler is the Domain Controller
# $LocalShare is the Exchange accessable folder where each backup is to be written to.
# $RemoteShare is the location of the offsite share where the script is to upload the fully completed PST backup file to.

[string] $Global:ExcDomainControler = (get-addomain).pdcemulator
[string] $Global:RemoteShare = #Path to external storage share (i.e. Azure Cloud - Example: "\\something.contoso.com\Archive\PST\"
[string] $Global:LocalShare = #Path to local network share - Example: "\\Server\ShareName\Output\"

#Function used to print status of the export attempt and to delay copying the mailbox until the process has completed,
#this will also skip mailboxes which are in "completed" status.
function ExchangeWaitLoop {
$global:starttime = get-date
clear-host

    [string] $Global:PSTName = $EmployeeMailboxName + ".pst"
    [string] $Global:filepath = $Global:LocalShare + $EmployeeMailboxName + $PSTName
    $Global:InitialPST = Get-MailboxExportRequest -Mailbox $EmployeeMailboxName | Get-MailboxExportRequestStatistics
    $Global:ExportStatus = $InitialPST.Status
    $NTAccount = Get-ADuser $EmployeeMailboxName
    $Global:EmployeeName = $NTAccount.Name

    if ($RunningStatuses -notcontains $ExportStatus){
		$RemoteDir = Get-ChildItem $RemoteShare
		$RemotePST = $RemoteDir.Name
		if($RemotePST -notcontains $PSTName){
@"
Beginning PST Backup process for $EmployeeName

******************************************
"@
                New-MailboxExportRequest -DomainController $ExcDomainControler -mailbox $EmployeeMailboxName -FilePath $filepath
                Start-Sleep -s 900
				}
    }
    [Bool]$global:ProcessActive = $True


    while ($ProcessActive -eq $True){
        
        $Global:ProcessingQueue = Get-MailboxExportRequest -Mailbox $EmployeeMailboxName | Get-MailboxExportRequestStatistics 

        #Find mailbox status after running backup request 
        $Global:PercentComplete = $ProcessingQueue.PercentComplete
        $Global:ExportStatus = $ProcessingQueue.Status

		    if ($ExportStatus -eq "Completed"){
            #Test to check if the mailbox process is complete, ends loop by setting Process Active flag to false
		    clear-host
		    [Bool]$Global:ProcessActive = $False
			[Bool]$PSTExists = test-path $filepath
			$ProcessingQueue = Get-MailboxExportRequest -Mailbox $EmployeeMailboxName | Get-MailboxExportRequestStatistics 

			#Find mailbox status after running backup request 
			$PercentComplete = $ProcessingQueue.PercentComplete
			$ExportStatus = $ProcessingQueue.Status
				if($ExportStatus -eq "Completed"){
					if ($PSTExists -eq $True){
						"Mailbox copy for User $EmployeeMailboxName completed, copying to remote share."
                        Start-Sleep -s 25
						move-item $filepath $RemoteShare
						#clear-host
						}
				}
				Elseif($PSTExists -eq $False){
				"No PST for user $EmployeeMailboxName exists, skipping to next user"
				}
		    }
            Elseif($ExportStatus -eq "Failed"){

            $PSTCopyErr = "PST Process For user $EmployeeMailboxName failed"
            write-host $PSTCopyErr
            $loglocation = $LocalShare + "output.txt"
            $PSTCopyErr >> $loglocation
            "Press any key to continue"
            pause
            }
            Elseif($ExportStatus -eq $null){
            [Bool]$Global:ProcessActive = $False
@"
Mailbox backup for $EmployeeName already completed
Starting next mailbox copy
"@
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
@"
    Mailbox for $EmployeeName backed up
    Next mailbox backup in 5 seconds
"@
    Start-Sleep -s 5
    clear-host
}

foreach ($mailbox in $Mailboxes) {
$global:EmployeeMailboxName = $mailbox.alias
ExchangeWaitLoop
}
