###########################################################################################################
#
# CRIAÇÃO DE SITES SHAREPOINT
# Existe um csv já designado para este script "SitesSharePoint.csv", todos os campos são obrigatórios
# É preciso especificar qual o User Admin na variável $OwnerAdmin (por default está o 1rjalves)
# É possível a criação de 1 ou MAIS sites em simultâneo
# O site ficará por default com o version history a 5
#
# Caso o site já exista (mesmo URl) irá aparecer um erro.
# Caso os campos não estejam todos preenchidos irá aparecer um erro.
# No final do script irá aparecer uma mensagem com os sites que foram Sucedidos/Não Sucedidos
#
# O user Admin que executar este código tem de estar adicionado à Enterprice APP "PnP Management Shell" - App ID:  e fazer um user consent"
# Só admins podem ser membros desta App "PnP Management Shell"
#
# Caso não tenha o Modulo PnP instalado:
# Install-Module -Name "PnP.PowerShell" -RequiredVersion 1.12.0 -Force -AllowClobber
#
# PnP não funciona com a versão (version PnP PowerShell 2.1.1.) - caso tenha a 2.1.1, por favor instalar o 1.12.0
# ver a resolução do problema aqui: https://learn.microsoft.com/en-us/answers/questions/1196279/import-module-could-not-load-file-or-assembly-syst
#
############################################################################################################



Connect-SPOService -Url "https://(...)-admin.sharepoint.com/"
Connect-PnPOnline -Url "https://(...)-admin.sharepoint.com/" -Interactive

$save = @()
$count=1

try{
    ##Import DATA from excel
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        InitialDirectory = [Environment]::GetFolderPath('Desktop')
        Filter           = 'Excel (*.CSV)|*.CSV'
        }
    $Null = $FileBrowser.ShowDialog()
    $DATA= Import-Csv -Path $FileBrowser.FileName
    
    ##Variáveis estáveis
    #$quotaMB = 102400
    [long]$quotaGB=0
    [long]$quota=0
    $VersioningLimit=5
    $OwnerAdmin = ""


    foreach($verification in $DATA)
    {
        $verificationtitle = $verification.Title
        $verificationowner = $verification.Owner
        $verificationquota = $verification.QuotaGB

        if($verificationtitle -eq "" -or $verificationowner -eq "" -or $verificationquota -eq "")
        {
            write-host "Não foi possível prosseguir. Existem campos por preencher." -f RED
            Break myLabel 
        }
    }

    :myLabel foreach($line in $DATA)
    {
        #Recolha da informação do excel para as variáveis
        Write-Host "$count) Site a ser criado..." -ForegroundColor Yellow
        $title = $line.Title
        $titlejoin = $title -replace '[^a-zA-Z0-9]', ''
        $url = "https://(...).sharepoint.com/sites/$titlejoin"
        $owner = $line.Owner
        [long]$quota = $line.QuotaGB
        [long]$quotaGB = $quota * 1024



        try{
            ##criação do site
            New-SPOSite -Title $title -Url $url -Owner $owner -StorageQuota $quotaGB  -Template "SITEPAGEPUBLISHING#0"

            Write-Host "Site $url criado. Waiting..." -ForegroundColor Yellow

            #Adicionar o user 1rjalves como admin do site
            Set-PnPTenantSite -Url $url -Owners $OwnerAdmin
            
            #Modificar o Histórico de versões para 5(default)
            Try{
                Connect-PnPOnline -Url $url -Interactive

                #Array to exclude system libraries
                $SystemLibraries = @("Form Templates", "Pages", "Preservation Hold Library","Site Assets", "Site Pages", "Images","Site Collection Documents", "Site Collection Images","Style Library","Recursos do Site","Páginas do Site")
         
                $Lists = Get-PnPList
        
                #Get All document libraries
                $DocumentLibraries = $Lists | Where {$_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $False -and $_.Title -notin $SystemLibraries}
    
                #Set Versioning Limits
                ForEach($Library in $DocumentLibraries)
                {
                    #powershell to set limit on version history
                    If($Library.EnableVersioning)
                    {
                        #Set versioning limit
                        Set-PnPList -Identity $Library -MajorVersions $VersioningLimit 
                        Set-PnPList -Identity $Library -MinorVersions $VersioningLimit 
                        Write-host -f Green "`tVersion History Settings has been Updated on '$($Library.Title)'"
                    }
                    Else
                    {
                        Write-host -f Yellow "`tVersion History is turned-off at '$($Library.Title)'"
                    }
                }
    
                Write-host -f Yellow "Site $url has been Updated"

                }Catch {
                    Write-host -f Red "Não foi possivel Alterar o número de versões para este site "$url
                    Write-host -f Red "Error:" $_.Exception.Message     
                }

                $count += $count
                
                #Guarda o site que foi BEM sucessido
                $save += [PSCUSTOMOBJECT] @{
                    "Status" = "OK";
                    "URL" = $url;
                    }
            }
        catch
        {
            Write-host "Não foi possivel criar este site $url, já pode existir com este nome ou mal preenchido" -ForegroundColor Red
            #Guarda o site que foi MAL sucessido
            $save += [PSCUSTOMOBJECT] @{
                "Status" = "NOT OK";
                "URL" = $url;
                }
        }
    }
}
catch{
    write-host "Não foi possível iniciar" -f RED
}

#Faz a listagem dos sites criados/não criados
Write-Host ""
Write-Host "-------------------------------------------------------------------------------"
Write-Host "Sites Criados" -ForegroundColor Green
$save