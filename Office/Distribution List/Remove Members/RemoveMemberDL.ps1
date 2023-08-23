############################################################################################
#
# Remover Membros de uma Distribution List
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

    #Validar se Endereço existe
    try{
        $validation = Get-DistributionGroup -Identity $DL -ErrorAction Stop
    }
    catch
    {
        Write-Host "O Endereço colocado NÃO EXISTE" -f Red
        Pause
        Break
    }

    #ADICIOANR TODOS MEMBROS A UM ARRAY $LISTA
    foreach($List in $DATA.1)
    {
        $a = $lista.add($List)
    }

    #REMOVER USERS À DL
    foreach($users in $lista)
    {
        try{
            Remove-DistributionGroupMember -Identity $DL -Member $users -Confirm:$False -ErrorAction Stop
            $i=$i+1
            Write-host "$i - $users REMOVIDO" -f Green
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
    Write-host "Os seguintes users NÃO FORAM REMOVIDOS" -f Red
    Write-host "MOTIVO: Utilizadores já Não se encontram no Grupo ou o endereço está errado... VALIDAR" -f Red
    $save
}
else
{
    Write-host $i "Utilizadores Removidos" -f Green
}
