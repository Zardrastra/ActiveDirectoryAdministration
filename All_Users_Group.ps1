$Time = Get-Date
$DateRaw = Get-Date
[string]$DateString = $DateRaw.ToLongDateString()
[string]$TimeString = $DateRaw.ToLongTimeString()
$TimeString = $TimeString | ForEach-Object {
    $_ -replace ":", "."
}

$dir = $env:USERPROFILE + "\Reporting\Active Directory\Security Group Reports\" + $DateString
New-Item -Path "$dir" -Type directory | Out-Null

$EmployeeMemberOf = ADUser -Filter * -SearchBase "OU=Users,DC=lapsang,DC=local" -Properties *


foreach ($user in $EmployeeMemberOf){
$Account = $user.SamAccountName
$MembersGp = Get-ADPrincipalGroupMembership -Identity $Account
$csvout = $dir + "\" + $Account + ".csv"
$MembersGp | Export-Csv $csvout
}
