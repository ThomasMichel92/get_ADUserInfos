$dateString = Get-Date -Format "yyyy-MM-dd_HHmmss"
$filename = "UserInformations_$dateString.csv"
$outputFile = "C:\PowershellScripts\Output\UserInformations_$dateString.csv"


# Setting Imported Properties
$users = Get-ADUser -Filter * -Properties LastLogon, LastLogonTimeStamp, PwdLastSet, PasswordLastSet, mail, ProxyAddresses | Sort-Object Name
$ouGroupFilter = @("OU=*****", "OU=****")

$output = foreach ($user in $users) {
    #Get All Groups of the User
    $ou = (Get-ADOrganizationalUnit -LDAPFilter "(distinguishedName=$($user.DistinguishedName.Split(',')[1..$($user.DistinguishedName.Split(',').Length - 1)] -join ','))").DistinguishedName
    
    #Loop over every Group
    foreach ($ouGroup in $ouGroupFilter) {
    # Check if the OU group name is in the distinguished name of the user's OU
    if ($ou -like "*$ouGroup*") {

      # Add the user's name and the OU group name to the filtered results array
      $ouGroupName = $ouGroup -replace "^OU=",""
      $filteredResults = $ouGroupName
      break
    }
    else {
      $filteredResults = "other"  
    }
    }

    #Setting the Output
    [PSCustomObject]@{
        Name = $user.Name
        Standort = $filteredResults
        Mail = $user.mail
        ProxyMail = $user.ProxyAddresses.Value
        LastLogon = if ($user.LastLogon) { [DateTime]::FromFileTime($user.LastLogon) } else { $null }
        LastLogonTimestamp = if ($user.LastLogonTimeStamp) { [DateTime]::FromFileTime($user.LastLogonTimeStamp) } else { $null }
        PwdLastSet = $user.PwdLastSet
        PasswordLastSet = $user.PasswordLastSet
        
    }
}

$output | Export-Csv $outputFile -NoTypeInformation

#Uplaod to Nextcloud
$nc_url="https://****"
$nc_dir ="/IT-Themen/AD-Übersicht"
$username="****"
$pwd="****"

$fileBytes = [System.IO.File]::ReadAllBytes($outputFile)
$uploadUrl = "$nc_url/remote.php/dav/files/$username/$nc_dir/$filename"

Write-Output $uploadUrl

$headers = @{
    Authorization = "Basic "+[System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($username+":"+$pwd))
}
Invoke-WebRequest -Uri $uploadUrl -Headers $headers -Method PUT -InFile $outputFile