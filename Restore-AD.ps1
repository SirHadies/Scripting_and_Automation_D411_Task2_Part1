#Alec Farrens #00384632

#Check Finance OU
$ouExists = Get-ADOrganizationalUnit -Filter {Name -eq "Finance"} -ErrorAction SilentlyContinue

if ($ouExists) {
    Write-Host "'Finance' OU already Exists. Deletion and Recreation Occurring.... "
    remove-ADOrganizationalUnit -Identity 'OU=Finance,DC=Consultingfirm,DC=com' -confirm:$false
    Write-Host "The 'Finance' OU has been removed succesfully"
}

#Create Finance OU
write-Host "Creating 'Finance' OU...."
New-ADOrganizationalUnit -Name 'Finance' -path 'DC=consultingfirm,DC=com' -DisplayName 'Finance' -ProtectedFromAccidentalDeletion $false
Write-Host "The 'Finance' OU has been created"

#Import CSV Users into Finance OU Group
$NewAD = Import-Csv $PSScriptRoot\financePersonnel.csvc

foreach ($ADUser in $NewAD) {
    
    
    $First = $ADUser.First_Name
    $Last = $ADUser.Last_Name
    $Name = $First + " " + $Last
    $SamName = $ADUser.samAccount
    $Postal = $ADUser.PostalCode
    $Office = $ADUser.OfficePhone
    $Mobile = $ADUser.MobilePhone
    
New-AdUser -GivenName $First -Surname $Last -Name $Name -SamAccountName $SamName -DisplayName $Name -PostalCode $Postal -MobilePhone $Mobile -OfficePhone $Office -Path $newOU
    
    }
 Get-ADUser -Filter * -SearchBase “ou=Finance,dc=consultingfirm,dc=com” -Properties DisplayName,PostalCode,OfficePhone,MobilePhone Out-File -FilePath "$PSScriptRoot\AdResults.txt"


 ##possible fix
 # Import users from CSV
$NewAD = Import-Csv "$PSScriptRoot\financePersonnel.csv"
$newOU = "OU=Finance,DC=consultingfirm,DC=com"

# Create users in AD
foreach ($ADUser in $NewAD) {
    $First = $ADUser.First_Name
    $Last = $ADUser.Last_Name
    $Name = "$First $Last"
    $SamName = $ADUser.samAccount
    $Postal = $ADUser.PostalCode
    $Office = $ADUser.OfficePhone
    $Mobile = $ADUser.MobilePhone

    try {
        New-AdUser -GivenName $First -Surname $Last -Name $Name -SamAccountName $SamName `
        -DisplayName $Name -PostalCode $Postal -MobilePhone $Mobile -OfficePhone $Office -Path $newOU
    } catch {
        Write-Error "Failed to create user $Name: $_"
    }
}

# Export AD users to a text file
$outputPath = "$PSScriptRoot\AdResults.txt"
Get-ADUser -Filter * -SearchBase "OU=Finance,DC=consultingfirm,DC=com" -Properties DisplayName,PostalCode,OfficePhone,MobilePhone |
Select-Object DisplayName, PostalCode, OfficePhone, MobilePhone |
Out-File -FilePath $outputPath
Write-Host "AD user details exported to $outputPath."