$file_data = Get-Content C:\Users\ricardo.j.alves\Desktop\dados.txt
$save = @()
$count=0
foreach($user in $file_data)
{ 
    $LicensesNumero=0
    #Get informations from Azure (user)
    $Get = Get-AzureADUser -ObjectId $User
    #Get informations from Service (user)
    $GetMsolUser = Get-MsolUser -UserPrincipalName $User


    $Enabled=$Get.AccountEnabled

    $LicensesNumero = (Get-MsolUser -UserPrincipalName $user |Select-Object -ExpandProperty Licenses | Where-Object {$_.AccountSkuID -like "*:ENTERPRISEPACK" -or $_.AccountSkuID -like "*:EMS" -or $_.AccountSkuID -like "*:SPE_F1" -or $_.AccountSkuID -like "*:M365_F1_COMM"}).count

    if($LicensesNumero -lt 1)
    {
        $Licenses = "Nao"
    }
    else
    {
        $Licenses = "Sim"
    }


    #Guarda todas as variaveis neste array
    $save += [PSCUSTOMOBJECT] @{
                "Email" = $User;
                "Enable/Disable" = $Enabled;
                "Caixa de Email?" = $Licenses;
                }
                    
    $count ++
    $count

}

#exportar as imformações para um CSV
$save | Export-csv -Path "C:\temp\Caixa de email.csv" -NoTypeInformation

