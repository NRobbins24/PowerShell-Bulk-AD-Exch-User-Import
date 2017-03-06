## ROBBINS MASS USERS SCRIPT
# Supply the password for all the newly created accounts. Swap next two lines if per-batch password is desired
# $password = read-host "Enter Default Password for all accounts created" -AsSecureString
$password = ConvertTo-SecureString "PassWord1234!@#$" -AsPlainText –Force
$ou = ""

# Define these as you wish 
$csvPath = "C:\users.csv"
#$userDB = "EXCH_DB_A-F"
$logPath = "C:\log.txt"

# When you change the OU and associated variable names be sure to also change
# the cooresponding variables below in the OU selection code block
$dept1 = "example.domain/OU/Tenants/ORG/DEPT1"
$dept2 = "example.domain/OU/Tenants/ORG/DEPT2"
$dept3 = "example.domain/OU/Tenants/ORG/DEPT3"
$dept4 = "example.domain/OU/Tenants/ORG/DEPT4"
$dept5 = "example.domain/OU/Tenants/ORG/DEPT5"


# OU Selection Error Message
$errorMessage = "*** INVALID ENTRY ***"
Import-module ActiveDirectory
# OU Selection message
Write-Host -foregroundcolor Cyan "`nTF [REMOVED] Mass User Creator Application"
Write-Host -foregroundcolor Cyan "Powershell Script Developed by Robbins`n`n"
Write-Host -foregroundcolor Red "NOTE: Currently running this as $env:username . <- THIS SHOULD BE YOUR ADMIN ACCOUNT OR YOU MAY RECEIVE ERRORS!"

Write-Host -foregroundcolor Cyan "Processing the Following Users For New Accounts:"
Get-Content $csvPath | select -Skip 1

Write-Host -foregroundcolor Cyan "`n You Must Select An OU For These Users"
$selectionChoice = Write-Host -foregroundcolor Cyan "Choose from the following OUs: `n" `
"1 = DEPARTMENT 1 `n 2 = DEPARTMENT 2 `n 3 = DEPARTMENT 3 `n 4 = DEPARTMENT 4 `n 5 = DEPARTMENT 5`n" `

# OU input message
$ouSelection = Read-Host "Which OU"

# *** Start OU Selection Code Block

$ou = $ouSelection

if(!($ou)) {
 Throw " NO OPTION SELECTED "
 Exit
}

Switch($ou) {
 1 {$ou = $dept1}
 2 {$ou = $dept2}
 3 {$ou = $dept3}
 4 {$ou = $dept4}
 5 {$ou = $dept5}
 default {
 $ou = $dept1
Write-Host -foregroundcolor Red "`nDEFAULTING TO DEPT1 OU DUE TO ERROR!"
#removed code to echo error. Defaulting OU
 Exit
 }
}
# *** End OU Selection Code Block
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
Import-CSV "$csvPath" | ForEach {
 
#USERS CORRECT MAILBOX STORE CHECK
#Reimporting CSV. Terrible practice... I know.
Param([string]$fileName)
 $users = Import-Csv $csvPath
 foreach ($user in $users){
 $flln = $user.'last name'.ToUpper().Substring(0)
 $db = ""
 if(($flln.CompareTo("A") -ge 0) -and ($flln.CompareTo("F") -le 0)){
 $db = "EXCH_DB_A-F"
 }
 elseif($flln.CompareTo("G") -ge 0 -and $flln.CompareTo("M") -le 0){
 $db = "EXCH_DB_G-M"
 }
 elseif($flln.CompareTo("N") -ge 0 -and $flln.CompareTo("S") -le 0){
 $db = "EXCH_DB_N-S"
 }
 elseif($flln.CompareTo("T") -ge 0 -and $flln.CompareTo("Z") -le 0){
 $db = "EXCH_DB_T-Z"
 }
 else{
 $db = "EXCH_DB_T-Z"
 }
}
$displayName = ($_.'Last Name' + " " + $_.'Rank' + " " + $_.'first name' + " " + $_.'MI' + " " + "[REMOVED] " + $_.'Unit' + " " + $_.'Title')
$sam = ($_.'First Name' + "." + $_.'Last Name')
$upn = ($_.'First Name' + "." + $_.'Last Name' +"@example.domain")
$AccountEx = $_.'Expires'

New-Mailbox `
-verbose `
-Password $password `
-Name $displayName `
-Alias $sam `
-OrganizationalUnit $ou `
-sAMAccountName $sam `
-FirstName $_.'First Name'`
-LastName $_.'Last Name'`
-DisplayName $displayName `
-UserPrincipalName $upn `
-Database $db `
-ResetPasswordOnNextLogon $true

Set-ADAccountExpiration $sam -DateTime "$AccountEx"
 
} | out-file $logPath
Import-CSV $csvPath | % {
$user1 = ($_.'First Name' + "." + $_.'Last Name')
Add-ADGroupMember -Identity "AD SECURITY GROUP - ALL USERS" -Member $user1
}
Write-Host -foregroundcolor Cyan "SUCCESS! Users created. Information Below..."
Get-Content $logPath

Exit
