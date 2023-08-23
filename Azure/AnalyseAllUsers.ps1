########################################################################################
# Script vai criar um csv com as seguintes informações em relação a TODOS os users da AD:
# UPN, DisplayName, Email, UserType, enable/disable, data da criação ,Manager, Licenças e Verifica se é Shared Mailbox ou não
# Na variável $Path, escolher o caminho para onde vai ser gravado o ficheiro csv
#
#######################################################################################

Connect-ExchangeOnline
Connect-AzureAD
Connect-MsolService


$Path="C:\temp\AnaliseUsers.csv"
$count=0
$save = @()

#Get all users from AD
$users = Get-AzureADUser -All $True


Foreach($User in $users)
{
    
        $objectid = $User.ObjectId

        #UPN
        $UPN=$User.UserprincipalName

        #saber o Manager(owner)
        $owner= (get-AzureADUserManager -ObjectId $objectid).UserprincipalName

        #Get informations from Service (user)
        $GetMsolUser = Get-MsolUser -UserPrincipalName $UPN
        
        
        $UserType = $User.UserType

        #DisplayName
        $DisplayName=$User.DisplayName

        $email = $User.Mail

        #Tipo de sincronização (cloud ou on-prem)
        $SyncUser=$User.dirsyncenabled

        if ($SyncUser -eq "TRUE")
        {
            $Sync="ON-PREMISE"
        }
        else{
            $Sync="CLOUD"
        }

        #Quando a conta foi criada #
        $Created=$GetMsolUser.WhenCreated
    
        #A conta está ativa ou desativa?
        $Enabled=$User.AccountEnabled

	try{
		$confirmshared = Get-Mailbox -Identity $UPN -RecipientTypeDetails SharedMailbox -ErrorAction Stop
		$shared = "Yes"
	}
	catch
	{
		$shared = ""
	}


	#Verifica se o utilizador tem uma destas licenças (F1-F3-E3)
	$LicensesNumero = (Get-MsolUser -UserPrincipalName $UPN |Select-Object -ExpandProperty Licenses | Where-Object {$_.AccountSkuID -like "*:ENTERPRISEPACK" -or $_.AccountSkuID -like "*:EMS" -or $_.AccountSkuID -like "*:SPE_F1" -or $_.AccountSkuID -like "*:M365_F1_COMM"}).count

    	if($LicensesNumero -lt 1)
    	{
        	$Licenses = "No"
    	}
    	else
    	{
        	$Licenses = "Yes"
    	}
    
    #Guarda todas as variaveis neste array
    $save += [PSCUSTOMOBJECT] @{
                "UPN" = $UPN
                "DisplayName"=$DisplayName;
                "Email" = $email;
                "User Tye" = $UserType;
                "Enable" = $Enabled;
                "SYNC" = $Sync;
                "Created on"=$Created;
                "Owner" = $owner
		        "Shared Mail Box" = $shared
		        "Licencas" = $Licenses
                }

    $count ++
    $count
    }

#exportar as imformações para um CSV
$save | Export-csv -Path $Path -NoTypeInformation
