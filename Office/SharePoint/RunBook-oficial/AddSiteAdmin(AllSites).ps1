###########################################################################################################
#
# Adicionar o User Admin como *Site admin* a todos os sites SharePoint
# É preciso especificar qual o User Admin na variável $Owner
#
# O user Admin que executar este código tem de estar adicionado à Enterprice APP "PnP Management Shell" - App ID: 31359c7f-bd7e-475c-86db-fdb8c937548e e fazer um user consent"
# Só admins podem ser membros desta App "PnP Management Shell"
#
############################################################################################################


##################################################
#PARAMETERS TO NOT CHANGE
$TenantAdminURL = "https://worten-admin.sharepoint.com/"
$Exclude = "POINTPUBLISHINGTOPIC#0","POINTPUBLISHINGHUB#0", "TEAMCHANNEL#1", "TEAMCHANNEL#0", "RedirectSite#0","SPSMSITEHOST#0", "EDISC#0","STS#-1","SRCHCEN#0"
$countSite=1
##################################################

$Owner = "1rjalves@worten.pt"

#Connect to PnP Online
Connect-PnPOnline -Url $TenantAdminURL -Interactive

$SitesURL = Get-PnPTenantSite | ? {$_.Template -notin $Exclude}

Foreach($Site in $SitesURL)
{
    try{
        Set-PnPTenantSite -Url $Site.URL -Owners $Owner
        Write-Host -f Green "$countSite - $Owner Adicionado ao Site" $Site.URL
    }catch{
        Write-Host -f Red "$countSite - $Owner NÃO Adicionado ao Site" $Site.URL $_.Exception.Message
    }

    $countSite += 1
}