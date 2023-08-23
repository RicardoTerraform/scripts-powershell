########################################################################################
#
#    
#
#Connect-ExchangeOnline
#Connect-AzureAD
#Connect-MsolService

#aliexpress@worten.pt
$Path="C:\temp\AnaliseUsers.csv"
$count=0
$save = @()




#import data from excel
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter           = 'Excel (*.CSV)|*.CSV'
}
$Null = $FileBrowser.ShowDialog()

$DATA= Import-Csv -Path $FileBrowser.FileName





Foreach($User in $DATA.userPrincipalName)
{
    
    try{
        $validation = Get-AzureADUser -ObjectId $User -ErrorAction Stop
    
    if($validation -notlike $Null)
    {
        #Saber se a conta é uma Shared MailBox ou não
        try
        {
            $SMB=Get-Mailbox -RecipientTypeDetails SharedMailbox -Identity $User -ErrorAction Stop
            $Mailbox="YES"

            
        }
        catch
        {
            $Mailbox="NO"
        }


        #saber o Manager(owner) da SMB
        $owner= (get-AzureADUserManager -ObjectId $User).UserprincipalName


        #Get informations from Azure (user)
        #$Get = Get-AzureADUser -ObjectId $User
        #Get informations from Service (user)
        #$GetMsolUser = Get-MsolUser -UserPrincipalName $User
        $UserType = (Get-AzureADUser -ObjectId $User).UserType

        #DisplayName
        $DisplayName=(Get-AzureADUser -ObjectId $User).DisplayName
        #$DisplayName=$Get.DisplayName

        #UPN
        #$UPN=(Get-AzureADUser -ObjectId $User).UserPrincipalName


        #Tipo de sincronização (cloud ou on-prem)
        $SyncUser=(Get-AzureADUser -ObjectId $User).dirsyncenabled
        #$SyncUser=$Get.dirsyncenabled

        if ($SyncUser -eq "TRUE")
        {
            $Sync="ON-PREMISE"
        }
        else{
            $Sync="CLOUD"
        }

        #Quando a conta foi criada #
        $Created=(Get-MsolUser -UserPrincipalName $User).WhenCreated
        #$Created=$GetMsolUser.WhenCreated

        #Licenças#
        #Yes or NO
        $License=(Get-MsolUser -UserPrincipalName $User).isLicensed
        #$License=$GetMsolUser.isLicensed

        #Caso tenha Licenças, Quantas tem#
        $numero=(Get-MsolUser -UserPrincipalName $User).Licenses.AccountSkuId.count
        #$numero=$GetMsolUser.Licenses.AccountSkuId.count


        #Caso tenha Licenças, indica quais são
        #$Lic=(Get-MsolUser -UserPrincipalName $User).licenses.AccountSkuId -join ";"
        #$Lic=$GetMsolUser.licenses.AccountSkuId -join ";"



        $LicensesNumero = (Get-MsolUser -UserPrincipalName $User |Select-Object -ExpandProperty Licenses | Where-Object {$_.AccountSkuID -like "*:ENTERPRISEPACK" -or $_.AccountSkuID -like "*:EMS" -or $_.AccountSkuID -like "*:SPE_F1" -or $_.AccountSkuID -like "*:M365_F1_COMM"}).count

        if($LicensesNumero -lt 1)
        {
            $Licenses = "Nao"
            $licenseType = ""
        }
        else
        {
            $Licenses = "Sim"

            If((Get-MsolUser -UserPrincipalName $User).Licenses.AccountSkuId -like "*:ENTERPRISEPACK")
            {
                $licenseType = "E3"
            }

            If((Get-MsolUser -UserPrincipalName $User).Licenses.AccountSkuId -like "*:EMS")
            {
                $licenseType = "E3"
            }

            If((Get-MsolUser -UserPrincipalName $User).Licenses.AccountSkuId -like "*:SPE_F1")
            {
                $licenseType = "F3"
            }

            If((Get-MsolUser -UserPrincipalName $User).Licenses.AccountSkuId -like "*:M365_F1_COMM")
            {
                $licenseType = "F1"
            }
        }

    
        #A conta está ativa ou desativa?
        $Enabled=(Get-AzureADUser -ObjectId $User).AccountEnabled
        #$Enabled=$Get.AccountEnabled

        #Last sing-in
        #mailbox
        try
        {
            $LastSignInMail=((Get-MailboxStatistics $User -ErrorAction Stop).LastLogonTime).ToString("dd-MM-yyyy")
        }
        catch
        {
            $LastSignInMail=$Null
        }
        #Azure
        try
        {
            #$LastSignInAzure=((Get-AzureAdAuditSigninLogs -top 1 -Filter "userprincipalname eq '$User'").CreatedDateTime).Substring(0,10)
            [DATETIME]$LastSignInAzure=((Get-AzureAdAuditSigninLogs -top 1 -Filter "userprincipalname eq '$User'").CreatedDateTime).Substring(0,10)
            $LastSignInAzureV2=$LastSignInAzure.ToString("dd-MM-yyyy")    
        }
        catch
        {
            $LastSignInAzureV2=$Null
        }

        }
        else
        {
            $DisplayName = "DELETED"
            $UserType =$Null
            $Mailbox=$Null
            $Enabled=$Null
            $Created=$Null
            $License=$Null
            $numero=$Null
            $Licenses=$Null
            $licenseType=$Null
            $Sync=$Null
            $LastSignInAzureV2=$Null
            $LastSignInMail=$Null
            $owner=$Null
        }
        }
        catch
        {
            $DisplayName = "DELETED"
            $UserType =$Null
            $Mailbox=$Null
            $Enabled=$Null
            $Created=$Null
            $License=$Null
            $numero=$Null
            $Licenses=$Null
            $licenseType=$Null
            $Sync=$Null
            $LastSignInAzureV2=$Null
            $LastSignInMail=$Null
            $owner=$Null
        }
    
    #Guarda todas as variaveis neste array
    $save += [PSCUSTOMOBJECT] @{
                "Email" = $User;
                "DisplayName"=$DisplayName;
                "User Tye" = $UserType
                "Shared Mail Box ? " = $Mailbox;
                "Enable" = $Enabled
                "Created on"=$Created;
                "Licenciado" = $License;
                "N de Licencas" = $numero;
                "Caixa de Correio ?" = $Licenses;
                "Qual Licenca ?" = $licenseType;
                "SYNC" = $Sync;
                "Last Sign-in (Azure)" = $LastSignInAzureV2;
                "Last Sign-in (Mail)" = $LastSignInMail;
                "Owner SMB" = $owner
                }

    $count ++
    $count
    }

#exportar as imformações para um CSV
$save | Export-csv -Path $Path -NoTypeInformation
