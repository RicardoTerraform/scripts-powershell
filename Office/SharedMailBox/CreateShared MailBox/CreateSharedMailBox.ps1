Connect-ExchangeOnline –Credential $Credential

try
{
    $TestAzureADConnection = Get-AzureADCurrentSessionInfo -ErrorAction Stop
}
catch
{
    Connect-AzureAD 
}

try{
    #import data from excel
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        InitialDirectory = [Environment]::GetFolderPath('Desktop')
        Filter           = 'Excel (*.CSV)|*.CSV'
    }
    $Null = $FileBrowser.ShowDialog()


    $DATA= Import-Csv -Path $FileBrowser.FileName



    #VERIFICAR SE OS CAMPOS ESTÃO PREENCHIDOS
    if($DATA.DisplayName -eq "" -or $DATA.Email -eq "" -or $DATA.Owner -eq "")
    {
        write-host "Não foi possível criar a SHARED MAILBOX. Existem campos por preencher" -f RED
        Pause
        Break
    }

    #VERIFICAR SE O OWNER EXISTE
    try{
        $verificar = Get-mailbox -Identity $DATA.Owner -ErrorAction Stop
    }
    catch
    {
        write-host "O Owner NÃO EXISTE, Verificar campo antes da criação da Shared MailBox" -f RED
        Pause
        Break
     }

    #CREATE SHARED MAILBOX
    try{
        $name = ($DATA.Email)
        $Names= $name.Substring(0,$name.Length-10)
        New-MailBox -Shared -Name $Names -DisplayName $DATA.DisplayName -PrimarySmtpAddress $DATA.Email -ErrorAction Stop
    }
    catch
    {
        Write-host "Shared MailBox com o endereço " $DATA.Email " JÁ EXISTE" -f RED
        Pause
        Break
    }

    Write-Host "Waiting 30s to change the UPN" -f Yellow
    Start-Sleep -Seconds 30

    #Change UPN Azure AD
    $localuser = Get-AzureADUser -SearchString $DATA.Email
    $localuser | foreach {$newUpn = $_.UserPrincipalName.Replace("worten.onmicrosoft.com","worten.pt"); $_ | Set-AzureADUser -UserPrincipalName $newUpn}
    $localuserObjectID = (Get-AzureADUser -SearchString $DATA.Email).ObjectID
    
    #Change MANAGEBY Azure AD
    $OwnerObjectID = (Get-AzureADUser -SearchString $DATA.Owner).ObjectID
    Set-AzureADUserManager -ObjectId $localuserObjectID -RefObjectId $OwnerObjectID

    #ALTERAR PARAMETROS PARA GUARDAR COPIAS DOS EMAILS ENVIADOS
    Set-Mailbox $DATA.Email -MessageCopyForSentAsEnabled $true
    Set-Mailbox $DATA.Email -MessageCopyForSendOnBehalfEnabled $true

    Write-Host ""
    Write-Host "Shared MailBox Criada" -f Green
    Pause
}
catch
{

}
