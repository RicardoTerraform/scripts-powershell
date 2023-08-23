#########################################################################################################
#
# Tarefa: Cria para um ficheiro todos os elementos de um determinado grupo
# os ObjectIDs dos grupos devem estar num ficheiro .CSV
# Podem ser adicionados vários grupos ao ficheiro .CSV
# O script cria ficheiros CSV separados pelos grupos adicionados
# Se dentro do grupo existirem mais grupos esses Grupos também serão adicionados a um ficheiro separado
# Mudar a variável $path para o caminho que querem
#
#########################################################################################################


Connect-AzureAD
$path="C:\temp\" 
$Notsave=@()

try{
    
    #import data from excel
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter           = 'Excel (*.CSV)|*.CSV'}

    $Null = $FileBrowser.ShowDialog()

    $DATA= Import-Csv -Path $FileBrowser.FileName -Header 1
    $aux=@()
    foreach($select in $Data.1)
    {
        $aux += $select
    }

    for ($i=0; $i -lt $aux.Length; $i++) 
    {
        $GroupObjectID = $aux[$i]
        $save=@()
 
        try{
            #Vai buscar todos os membros associados ao grupo
            $members = Get-AzureADGroupMember -ObjectId "$GroupObjectID" -All $true                
                
            #Vai buscar o Id do grupo
            $GroupName = Get-AzureADGroup -ObjectId "$GroupObjectID" | Select-Object -ExpandProperty DisplayName

            #Preencher o array Save com todos os membros
            foreach($user in $members)
            {
                $save += [PSCUSTOMOBJECT] @{
                     "UPN"          = $user.UserPrincipalName
                     "Utilizadores" = $user.DisplayName
                     }   
            }

            foreach($user in $members)
            {
                $UserId = $user.ObjectId
                try{
                    $Grouptype = Get-AzureADGroup -ObjectId "$UserId" | Select-Object -ExpandProperty ObjectType      
                    $aux += $UserId
                }
                catch{}
             }               
                
             #Remove caracteres especiais do nome do grupo para utilizar como nome do ficheiro
             $String = $GroupName -replace '[^a-zA-Z0-9]', ''

             #cria a path do ficheiro com o nome do grupo
             $CompletePath = $path + $String + ".csv"

             #Cria o CSV 
             $save | Export-csv -Path $CompletePath -NoTypeInformation
             Write-host $GroupName " concluído" -f Green
        }
        catch
        {
            $Notsave += [PSCUSTOMOBJECT] @{
                "Grupo" = $GroupName
                }
        }
    }

}
catch{
Write-host "Alguma coisa falhou com a importação do ficheiro"
}

if(($Notsave).count -gt 0)
{
    write-host "Ficheiros dos seguintes grupos que não foram criados" -f Red
    $Notsave
}

Pause