Clear-Host
$host.ui.RawUI.WindowTitle = "Group Reporting Utility"

#Grab current time and set the output directory for the raw files
$Time = Get-Date
$DateRaw = Get-Date
[string]$DateString = $DateRaw.ToLongDateString()
[string]$TimeString = $DateRaw.ToLongTimeString()
$TimeString = $TimeString | ForEach-Object {
    $_ -replace ":", "."
}

$dir = $env:USERPROFILE + "\Reporting\Active Directory\Group Reports\" + $DateString + " " + $TimeString 


if (Get-Command dsquery -errorAction SilentlyContinue){

@"

Type in the name of the group you wish to check and press enter

"@

$global:GroupName = Read-Host 'Group name'
clear-host

"Report Running, please wait a moment for this to complete"

#create the new directory for the report and run
New-Item -Path "$dir" -Type directory | Out-Null
$GroupContent = Get-ADGroupMember $GroupName -Recursive| get-aduser  -Properties DisplayName, Department, Created, whenChanged, SamAccountName, PasswordLastSet, PasswordNeverExpires, PasswordNotRequired, CanonicalName, CN
$csvout = $dir + "\" + $GroupName + ".csv"
$htmout = $dir + "\" + $GroupName + ".htm"
$GroupContent | Export-Csv -Encoding unicode $csvout

#HTML Table Format
$Header = @"
<style>
BODY{background-color:peachpuff;}
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color:thistle}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color:palegoldenrod}
</style>
"@

#import CSV file
$AccountDataRaw = import-csv $csvout
ForEach ($account in $AccountDataRaw) {
$account.CanonicalName = $account.CanonicalName.TrimEnd($account.CN)
}
$AccountDataRaw | Export-Csv -Encoding unicode $csvout

#convert to HTML, clean up formatting and output to file
$AccountDataRaw | Select-Object @{Name = "Name"; Expression = {$_.CN}}, @{Name = "Active Account"; Expression = {$_.Enabled}}, @{Name = "Password Last Set"; Expression = {$_.PasswordLastSet}}, @{Name = "Password Never Expires"; Expression = {$_.PasswordNeverExpires}}, @{Name = "Department/AD Location"; Expression = {$_.CanonicalName}} | Sort-Object "Department/AD Location" | ConvertTo-HTML -Head $Header | Out-File $htmout

#remove the first line in the report and commit changes
Rename-Item $csvout  source.txt
$sourceloc = $dir + "\" + "source.txt"
$RAWCSV = $dir + "\" + $GroupName + "_Raw_Data.csv"
Get-Content $sourceloc | Select-Object -Skip 1 | Out-File -Encoding unicode $RAWCSV
Remove-Item $sourceloc
clear-host
@"

Report is complete and has been output in the folder:

"@
write-host $dir 
@"

_____________________________________________________________________________________________________
"@

invoke-item $dir
pause
}
else {
clear-host
$host.ui.RawUI.WindowTitle = "Error: Remote Server Administration Tools Missing"
@"
Error: Remote Server Administration tools for Windows not found.

Press Enter to exit the application, a window will then open where you can download and install the
Remote Server Administration Tools Update for your version of Windows


Download the package for your Windows version and architecture type (x86 or x64)

Install this package and relaunch the script, you will not need to reboot your computer

____________________________________________________________________________________________________
"@
pause
$InternetExplorer=new-object -com internetexplorer.application
$InternetExplorer.navigate2("https://support.microsoft.com/en-ie/help/2693643/remote-server-administration-tools-rsat-for-windows-operating-systems")
$InternetExplorer.visible=$true
exit
}
