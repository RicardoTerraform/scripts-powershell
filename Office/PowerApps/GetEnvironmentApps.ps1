$save = @()
$Path="C:\temp\PowerApps.csv"
$getpowerapps = Get-AdminPowerApp

foreach($getpwapp in $getpowerapps)
{
    
    $AppEnviroment= (Get-AdminPowerAppEnvironment -EnvironmentName $getpwapp.EnvironmentName).DisplayName
    $AppPowerapp = $getpwapp.DisplayName
    $CreatedTime= ($getpwapp.CreatedTime).substring(0,10)
    $Owner = $getpwapp.owner.displayName

    $save += [PSCUSTOMOBJECT] @{
                "AppEnvironment" = $AppEnviroment;
                "DisplayName" = $AppPowerapp;
                "CreatedTime"=$CreatedTime;
                "Owner" = $Owner
                }

}

$save | Export-csv -Path $Path -NoTypeInformation