#Environment Variables
$AppData = (Get-Item env:appdata).value
#For Outlook signatures
$sigPath = '\Microsoft\Signatures'
$localSignatureFolder = $AppData+$sigPath
$templateFilePath = "\\server\signatures"

$userName = $env:username

$filter = "(&(objectCategory=User)(samAccountName=$userName))"
$searcher = New-Object System.DirectoryServices.DirectorySearcher
$searcher.Filter = $filter
$ADUserPath = $searcher.FindOne()
$ADUser = $ADUserPath.GetDirectoryEntry()
$name = $ADUser.DisplayName
$email = $ADUser.mail
$department = $ADUser.department
$phone = $ADUser.telephonenumber
$office = $ADUser.physicalDeliveryOfficeName

$namePlaceHolder = "userDisplayName"
$emailPlaceHolder = "userEmail"
$departmentPlaceHolder = "userDepartment"
$phonePlaceHolder = "userTelephone"


$rawTemplate = get-content $templateFilePath"\assinatura.txt"

$signature = $rawTemplate -replace $namePlaceHolder,$name
$rawTemplate = $signature

$signature = $rawTemplate -replace $emailPlaceHolder,$email
$rawTemplate = $signature

$signature = $rawTemplate -replace $phonePlaceHolder,$phone
$rawTemplate = $signature

$signature = $rawTemplate -replace $departmentPlaceHolder,$department

$fileName = $localSignatureFolder + "\" + $userName + ".htm"

#Gets the template last update time
if(test-path $templateFilePath){
	$templateLastModifiedDate = [datetime](Get-ItemProperty -Path $templateFilePath -Name LastWriteTime).lastwritetime
}

# Checks if there is a signature and its last update time
if(test-path $filename){
	$signatureLastModifiedDate = [datetime](Get-ItemProperty -Path $filename -Name LastWriteTime).lastwritetime
}else{
	$signature > $fileName
	break
}


if((get-date $templateLastModifiedDate) -gt (get-date $signatureLastModifiedDate)){
	$signature > $fileName
}