# Variables
$SMTPServer = "mailrelay.lapsang.test"
$From = "computer-login-report@lapsang.test"
$ToSend = "employee.1@lapsang.test"
$MailSubject = "Computers logged in in past 30 days"


#Email Properties
$MailProperties = @{
 From = $From
 To = $ToSend
 Subject = $MailSubject
 SMTPServer = $SMTPServer
}

Out-String

$DateCutOff=(Get-Date).AddDays(-30)

$AllComputers = get-adcomputer -SearchBase 'DC=lapsang,DC=test' -Filter 'enabled -eq $true' -Property CN, CanonicalName, IPv4Address, IPv6Address, LastLogonDate | Where {$_.LastLogonDate -gt $datecutoff} | Select CN, CanonicalName, IPv4Address, IPv6Address, LastLogonDate | ConvertTo-Html
$emailBody = $TestGroup | Out-String

Send-MailMessage @MailProperties -body $emailBody -BodyAsHtml
