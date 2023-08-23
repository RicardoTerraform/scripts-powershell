###########################################################################################################
#
# Apagar Reclyc BIN (mover todos os items com mais de 15* dias para o second stage, e depois no second stage apaga tudo)
# O user Admin que executar este código tem de estar adicionado à Enterprice APP "PnP Management Shell" - App ID: 31359c7f-bd7e-475c-86db-fdb8c937548e e fazer um user consent"
# Só admins podem ser membros desta App "PnP Management Shell"
#
############################################################################################################

##################################################
#PARAMETERS TO NOT CHANGE
$TenantAdminURL = "https://worten-admin.sharepoint.com/"
$Exclude = "POINTPUBLISHINGTOPIC#0","POINTPUBLISHINGHUB#0", "TEAMCHANNEL#1", "TEAMCHANNEL#0", "RedirectSite#0","SPSMSITEHOST#0", "EDISC#0","STS#-1","SRCHCEN#0"

$i=1
$countDeletedItemTotal = 0
$countDeletedItemPerSite=0 
$CountItemSecondStage = 0
##################################################
$dias = 10

#Connect
Connect-PnPOnline -Url $TenantAdminURL -Interactive

#Get all Sites sharepoint
$SitesURL = Get-PnPTenantSite | ? {$_.Template -notin $Exclude}


foreach($site in $SitesURL){
    
    Connect-PnPOnline -Url $site.URL -Interactive
    Write-Host "$i) Eliminar ficheiros (Recycle bin) SharePoint do site" $site.URL
    $DateToday = Get-Date
    $FileDate = Get-PnPRecycleBinItem -FirstStage
    $countDeletedItemPerSite=0
     
    try{
        foreach ($file in $FileDate)
        {
            $getFileDate = $file.DeletedDate
    
            $Date = $DateToday - $getFileDate
            $DateFinal=$Date.Days
        
            if($DateFinal -gt $dias)
            {
	            $countDeletedItemPerSite = $countDeletedItemPerSite + 1
                #Write-Host "Vai ser movido para second stage"
                $ficheiro = $file.Id
                Get-PnPRecycleBinItem -Identity "$ficheiro" | Clear-PnpRecycleBinItem -Force
            }
        }
    }catch
    {}

    #Saber sumatório dos Items eliminados
    $countDeletedItemTotal =  $countDeletedItemTotal + $countDeletedItemPerSite
    Write-Host "Ficheiro com mais de $dias dias eliminados do First Stage - $countDeletedItemPerSite Eliminados"

    try{
        #Elimina todos os ficheiros do SECOND STAGE
        Get-PnPRecycleBinItem -SecondStage | Clear-PnpRecycleBinItem -Force
    }
    catch
    {}

    Write-Host "Waiting..."
    Write-Host "Todos os Ficheiros eliminados do Second Stage"
    Write-Host "Site" $site.URL "DONE!!!" -ForegroundColor Green
    Write-Host ""

    $i=$i+1
}

# Condição que vai validar se todos os Items do Second Stage (de cada site) foram eliminados
foreach($site in $SitesURL)
{    
    Connect-PnPOnline -Url $site.URL -Interactive

    $CountItemSecondStage = (Get-PnPRecycleBinItem -SecondStage).count
    
    if($CountItemSecondStage -gt 0)
    {
        Get-PnPRecycleBinItem -SecondStage | Clear-PnpRecycleBinItem -Force
        #Start-Sleep -Seconds 30
        Write-Host "Eliminou mais ficheiros do site "$site.URL -ForegroundColor red
    }
    if($CountItemSecondStage -eq 0)
    {
        Write-Host "DOUBLE CHECK (SECOND STAGE) - "$site.URL "- OK"
    }
}

Write-Host "Foram eliminados $countDeletedItemTotal items"