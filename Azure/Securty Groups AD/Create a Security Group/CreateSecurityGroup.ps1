##################################################################################################
#   Script vai correr um .TXT e vai criar Security groups na nossa AD com os nomes indicados no .TXT
#   Adicioanr os nomes (Display Name) que queremos para os grupos num .TXT
#   Na variável $Path adicionar o caminho para o .TXT indicado
#
##################################################################################################


$save = @()

$Path = "C:\Users\lista.txt"
$list = Get-Content $Path

$i=0


foreach($DL in $list)
{
    Write-host $DL -f Green
    $i=0
    New-AzureADGroup -DisplayName $DL -SecurityEnabled $true -MailEnabled $false -MailNickName $DL

 }

 Write-host "DONE" -f Green
 $save




