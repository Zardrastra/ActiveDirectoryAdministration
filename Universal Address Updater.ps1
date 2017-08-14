#Grab name of template user to be used as reference for Address update within sub OUs within ActiveDirectory
#Useful for batch updating child accounts under an OU where a template file exists
#Must be run using a Domain Administrator account 

function AddressUpdate {

#Input name of NT account where addresses are to be copied from
$employeeTemplate = Read-Host 'Employee Template to Apply'

#Create replication master copy object - This is the information that is to be updated on the Employee records in AD
$copyObject = get-aduser -identity $employeeTemplate -Properties Country,City,PostalCode,physicalDeliveryOfficeName,StreetAddress,State,wWWHomePage,Company

#This sets the organizational unit where the addresses are to be updated, the script uses the folder of the template as a point of reference
#For example, if the NT Address template user is in /company.com/Employees/Europe/UK/London_Office/
#All users in the London_Office will have their addresses updated to match the template as they are also under the London_Office
$AccountOUPath = Get-ADUser $employeeTemplate -Properties distinguishedname,cn | Select-Object @{n='ParentContainer';e={$_.distinguishedname -replace "CN=$($_.cn),",''}}

#This creates an object contianing all employees under the same Organizational Unit as the template user
$CurrentADUsers = Get-ADuser -Filter * -SearchBase $AccountOUPath.ParentContainer -properties Country,City,PostalCode,physicalDeliveryOfficeName,StreetAddress,State,wWWHomePage,Company

try {
#For each user in the CurrentADUsers apply the address information of the template
ForEach ($usrObj in $CurrentADUsers) { 
$usrObj.Country = $copyObject.Country
$usrObj.City = $copyObject.City
$usrObj.PostalCode= $copyObject.PostalCode
$usrObj.physicalDeliveryOfficeName = $copyObject.physicalDeliveryOfficeName
$usrObj.StreetAddress = $copyObject.StreetAddress
$usrObj.State = $copyObject.State
$usrObj.wWWHomePage = $copyObject.wWWHomePage
$usrObj.Company = $copyObject.Company

#Once the record has had the changes queued apply the pending changes 
Set-ADUser -instance $usrObj -ErrorVariable $ErrorAC
}
}
catch {
"There are issues with the details provided, kindly review the following error"
write-host $ErrorAC
Pause
AddressUpdate
}
}

#Invokes Address Updating Function 
AddressUpdate
