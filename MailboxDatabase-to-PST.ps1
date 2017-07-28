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
#this will also skip mailboxes which are in "completed" or failed status (In the case of failure it will write to a log and skip).
function ExchangeWaitLoop {
$Global:starttime = get-date
$Global:MailboxTotal = $Mailboxes.count
$Global:currentMailboxCount++
$TotalPercentComplete = ($currentMailboxCount/$MailboxTotal).ToString("P")
#Set a flag for mailbox download request false to allow the mailbox to be skipped if it is already backed up
[bool]$Global:propagationLock = $False
clear-host
    #Generate standard naming format for PST backup file - (Alias).PST
    [string] $Global:PSTName = $EmployeeMailboxName + ".pst"
    #Generate the path variable for where the PST is to be downloaded to temporarily
    [string] $Global:filepath = $Global:LocalShare + $PSTName
    $Global:InitialPST = Get-MailboxExportRequest -Mailbox $EmployeeMailboxName | Get-MailboxExportRequestStatistics
    $Global:ExportStatus = $InitialPST.Status

    #TODO: Resolve issue with NTACCOUNT naming to use other measure to find the user account name, this currently fails if the mailbox Alias is not the same as the employee NT account name
    $NTAccount = Get-ADuser $EmployeeMailboxName
    $Global:EmployeeName = $NTAccount.Name
#Check to test if request to backup mailbox to PST is to be sent to exchange server processing queue
    if ($RunningStatuses -notcontains $ExportStatus){
        #Check contents of remote share and test if mailbox.pst file for use already exists (in the form Alias.pst)
		$RemoteDir = Get-ChildItem $RemoteShare
        $RemotePST = $RemoteDir.Name
        #If the remote server does not have a file with the alias.pst for this employee add the mailbox to the current queue
		if($RemotePST -notcontains $PSTName){
@"
Beginning PST Backup process for $EmployeeName
******************************************
"@
                New-MailboxExportRequest -DomainController $ExcDomainControler -mailbox $EmployeeMailboxName -FilePath $filepath
                #Set a flag for mailbox propagation to prevent process active loop from exiting before the queue begins processing
				[bool]$Global:propagationLock = $True
					}
                
				}
	#Flag to set the process loop active, until false the while-loop will run
    [Bool]$Global:ProcessActive = $True


while ($ProcessActive -eq $True){
        #Get the current status of the mailbox in the exchange backup loop on each run of the loop.
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
						Try{
						"Mailbox copy for User $EmployeeMailboxName completed, copying to remote share."
                        Start-Sleep -s 25
						move-item $filepath $RemoteShare
						#clear-host
						}
						catch{
						currentTimeFN
						$PSTCopyErr = "$CurrentTimetext PST Copy process for user $EmployeeMailboxName failed"
						write-host $PSTCopyErr
						$loglocation = $LocalShare + "output.txt"
                        $PSTCopyErr >> $loglocation
                        pause
						}
						}
				}
				Elseif($PSTExists -eq $False){
				"No PST for user $EmployeeMailboxName exists, skipping to next user"
				}
		    }
            Elseif($ExportStatus -eq "Failed"){
			currentTimeFN
            $PSTCopyErr = "$CurrentTimetext PST Process For user $EmployeeMailboxName failed"
            write-host $PSTCopyErr
            $loglocation = $LocalShare + "output.txt"
            $PSTCopyErr >> $loglocation
            [Bool]$Global:ProcessActive = $False
            }
            #Nest elseif statement into a flag statement - Account Processing on/off
            Elseif(($ExportStatus -eq $null) -and($propagationLock -eq $False)){
					[Bool]$Global:ProcessActive = $False
@"
Mailbox backup for $EmployeeName already completed
Starting next mailbox copy
"@
            }

		    else {
#Else function to generate status message of the ongoing backup with timestamps
            Start-Sleep -s 5
			currentTimeFN
		    clear-host
if ($PercentComplete -ne $null){
$PercentCompleteText = "Percent completed: $PercentComplete %"
}
elseif ($PercentComplete -eq $null){
$PercentCompleteText = "Awaiting Mailbox Backup"
}
@"
Awaiting mailbox backup for:
$EmployeeName
Mailbox $currentMailboxCount of $MailboxTotal ($TotalPercentComplete Overall)
Start time: $starttime
$CurrentTimetext
$PercentCompleteText
Refreshing backup progress in 2 minutes
"@
Write-Host `a
Start-Sleep -s 120
clear-host
"Refreshing backup progress..."
		}
	}
	currentTimeFN
clear-host

@"
	$CurrentTimetext
    Mailbox $currentMailboxCount of $MailboxTotal ($TotalPercentComplete Overall)
    Mailbox for $EmployeeName backed up
    Next mailbox backup in 5 seconds
"@
    Start-Sleep -s 5
    clear-host
}


function currentTimeFN{
$Global:currentTime = get-date
$Global:CurrentTimetext = "Current time:" + " " + $currentTime.ToShortTimeString()
}
foreach ($mailbox in $Mailboxes) {
$global:EmployeeMailboxName = $mailbox.alias
ExchangeWaitLoop
}
