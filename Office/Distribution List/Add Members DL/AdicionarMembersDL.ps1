#
# Adicionar utilizadores a uma DL
# Não é preciso nenhum template, basta abrir um Documento excel (CSV) e adicionar a lista de utilizadores a começar na Coluna A1.
# Script consegue apanhar algumas exceções: caso o email da DL não exista, ou caso alguns dos membros listados não existam ou já tenham sido adicionados anteriormente. 
# Caso isto aconteça no final do script irá aparecer uma lista dos utilizadores que não foram adicionados.
#
############################################################################################

Connect-ExchangeOnline

$save = @()
$a=New-Object System.Collections.ArrayList
$i=0
$lista= New-Object System.Collections.ArrayList
try{
    
    #import data from excel
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        InitialDirectory = [Environment]::GetFolderPath('Desktop')
        Filter           = 'Excel (*.CSV)|*.CSV'
    }
    $Null = $FileBrowser.ShowDialog()


    $DATA= Import-Csv -Path $FileBrowser.FileName -Header 1
    $DL = Read-Host "Qual é o ENDEREÇO da Distribution List?"

    #ADICIOANR TODOS MEMBROS A UM ARRAY $LISTA
    foreach($List in $DATA.1)
    {
        $a = $lista.add($List)
    }

    #ADICIONAR USERS À DL
    foreach($users in $lista)
    {
        try{
            Add-DistributionGroupMember -Identity $DL -Member $users -ErrorAction Stop
            $i=$i+1
            Write-host "$i - $users DONE" -f Green
        }
        Catch
        {        
            $save += New-Object PSObject -Property @{
                "User"=$users}
        }
    }
}
catch
{

}

if($save.count -gt 0)
{
    Write-host ""
    Write-host "Os seguintes users NÃO FORAM ADICIONADOS" -f Red
    Write-host "MOTIVO: Utilizadores já foram adicionados ao Grupo ou Não existem... VALIDAR" -f Red
    $save
}
else
{
    Write-host ""
    Write-host $i "Utilizadores adicionados" -f Green
}

#Pause