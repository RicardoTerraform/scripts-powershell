##################################################################################
# Script vai alterar o estado de todas as contas adicionadas para DISABLE
# Lista de users devem ser adicioandos a um .TXT
# No atributo $Path inidicar o caminho para o .TXT 
#
#################################################################################

Connect-AzureAD
$Path = "C:\Users\dados.txt"
$file_data = Get-Content $Path
$save = @()
$count=0
foreach($shared in $file_data)
{ 
    Set-AzureADUser -ObjectId $shared -AccountEnabled False
}

Write-Host "completo" -f Green