###############################################################################################################
#  Script vai correr um .TXT e vai adicionar todos os users que estão no .TXT a um Security Groups na AD
#  Adicioanr os users que queremos num .TXT
#  Na variável $Path adicionar o caminho para o .TXT indicado
#  Na variável $ObjectidGroup adicionar o objectID, podemos ver esta informação do grupo da AD
#
################################################################################################################

Connect-AzureAD

$Path = "C:\Users\dados.TXT"
$file_path = Get-content -path $Path

$ObjectidGroup = "21a8e30f-8559-4706-9992-041cb150d062"

foreach($file in $file_path)
{
    $ob = Get-AzureADUser -SearchString $file | Select-Object -ExpandProperty Objectid
    Add-AzureADGroupMember -ObjectId $ObjectidGroup -RefObjectId $ob
}

Write-Host "DONE"