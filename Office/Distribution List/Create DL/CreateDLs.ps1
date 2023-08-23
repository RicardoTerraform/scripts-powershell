##########################################################################
#
# Criar uma Distribution List
# Junto do script está um csv (Template-Create-DL.csv) onde temos de preencher pelo menos 3 campos (Displayname, Email e Owners) *owners APENAS UM.
# Script consegue apanhar algumas exceções: caso o email da DL já exista no nosso 365, caso os campos indicados não sejam preenchidos ou caso algum Owner esteja mal preenchido ou não exista.
#
##########################################################################


Connect-ExchangeOnline
$Criada=@()
$NaoCriada=@()
try{
    #import data from excel
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        InitialDirectory = [Environment]::GetFolderPath('Desktop')
        Filter           = 'Excel (*.CSV)|*.CSV'
    }
    $Null = $FileBrowser.ShowDialog()
    $DATA= Import-Csv -Path $FileBrowser.FileName



    #VERIFICAR SE OS CAMPOS ESTÃO PREENCHIDOS, "description" não é obrigatório
    foreach($verification in $DATA)
    {
        $verificationDisplayName = $verification.DisplayName
        $verificationEmail = $verification.Email
        $verificationDescription = $verification.Description
        $verificationowner = $verification.Owner

        if($verificationDisplayName -eq "" -or $verificationEmail -eq "" -or $verificationowner -eq "")
        {
            write-host "Não foi possível prosseguir. Existem campos por preencher." -f RED
            Break myLabel 
        }
    }

    #VERIFICAR SE OS OWNERS EXISTEM ANTES DA CRIAÇÃO DA DL
    foreach($Owner in $DATA.Owner)
    {
        try{
            $verificar = Get-mailbox -Identity $Owner -ErrorAction Stop
        }
        catch
        {
            write-host "Um dos owners NÃO EXISTE, Verificar campo antes da criação da DL" -f RED
            Break myLabel 
        }
    }


    #VERIFICAR SE DISPLAYNAME DA DLS JÁ EXISTEM ANTES DA CRIAÇÃo
    :myLabel foreach($line in $DATA)
    {
        try{
            New-DistributionGroup -Name $line.DisplayName -DisplayName $line.DisplayName -PrimarySmtpAddress $line.Email -ErrorAction Stop

            if($line.Description -eq "")
            {
                Set-DistributionGroup $line.Email -ManagedBy $line.Owner -ErrorAction Stop
            }
            else{
                Set-DistributionGroup $line.Email -ManagedBy $line.Owner -Description $line.Description -ErrorAction Stop
            }
            
            $Criada += [PSCUSTOMOBJECT] @{
                "Displayname" = $line.DisplayName 
                }
               
        }
        catch
        {
            Write-host "Distribution List com o endereço " $line.Email " JÁ EXISTE" -f RED
            $NaoCriada += [PSCUSTOMOBJECT] @{
                "Displayname" = $line.DisplayName 
                }
        }
    }
}catch
    {
        write-host "Não foi possível criar a(s) DL(s), rever os parâmetros preenchidos" -f RED
    }

Write-host ""
Write-host ""
Write-host "Distribution List Criada:" $Criada  -f Green
Write-host "Distribution List Não Criada:" $NaoCriada -f RED