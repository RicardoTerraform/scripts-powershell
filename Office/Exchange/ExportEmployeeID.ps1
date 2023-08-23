$users = Get-AzureADUser -All $True 
$save=@()
foreach($user in $users)
{
    $displayname= $user.displayname
    $UPN = $user.UserPrincipalName
    $employeeID= $user.ExtensionProperty.employeeId

    $save += [PSCUSTOMOBJECT] @{
                "DisplayName" = $displayname;
                "UPN" = $UPN;
                "EmployeeID"=$employeeID;
                }

}

$save | Export-csv -Path "C:\temp\empIDv2.csv" -NoTypeInformation