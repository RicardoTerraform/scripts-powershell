########################################################################################
#
#    
#



#aliexpress@worten.pt
$Path="C:\temp\AnaliseSharedMailBox.csv"
$count=0
$save = @()
#$User="aliexpress@worten.pt"



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
    #Saber se a conta é uma Shared MailBox ou não
    try
    {
        $SMB=Get-Mailbox -RecipientTypeDetails SharedMailbox -Identity $User -ErrorAction Stop
        $Mailbox="Shared Mailbox"

        #saber o Manager(owner) da SMB
        $owner= (get-AzureADUserManager -ObjectId $User).UserprincipalName

        $DisplayName=(Get-AzureADUser -ObjectId $User).DisplayName

        $SyncUser=(Get-AzureADUser -ObjectId $User).dirsyncenabled
        if ($SyncUser -eq "TRUE")
        {
            $Sync="ON-PREMISE"
        }
        else{
            $Sync="CLOUD"
        }

        $Created=(Get-MsolUser -UserPrincipalName $User).WhenCreated



        #$License=(Get-MsolUser -UserPrincipalName $User).isLicensed



        #$numero=(Get-MsolUser -UserPrincipalName $User).Licenses.AccountSkuId.count

        #$Lic=(Get-MsolUser -UserPrincipalName $User).licenses.AccountSkuId -join ";"
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





        $Enabled=(Get-AzureADUser -ObjectId $User).AccountEnabled

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


            #Guarda todas as variaveis neste array
        $save += [PSCUSTOMOBJECT] @{
                "Email" = $User;
                "Shared/User" = $Mailbox;
                "DisplayName"=$DisplayName;
                "Enable" = $Enabled
                "Licenciado" = $Licenses;
                "License Type" = $licenseType
                "SYNC" = $Sync;
                "Created on"=$Created;
                "Last Sign-in (Azure)" = $LastSignInAzureV2;
                "Last Sign-in (Mail)" = $LastSignInMail;
                "Owner SMB" = $owner
                }

        $count ++
        $count
    

    }
    catch
    {
    }

   
 }   



#exportar as imformações para um CSV
$save | Export-csv -Path $Path -NoTypeInformation
