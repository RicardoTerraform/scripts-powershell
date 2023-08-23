###########################################################################################################
#
# Adicionar o User Admin como *Site admin* a um Site específico
# É preciso especificar qual o User Admin na variável $Owner
# É preciso especificar qual o site sharepoint na variável $SiteURL
#
# O user Admin que executar este código tem de estar adicionado à Enterprice APP "PnP Management Shell" - App ID: 31359c7f-bd7e-475c-86db-fdb8c937548e e fazer um user consent"
# Só admins podem ser membros desta App "PnP Management Shell"
#
############################################################################################################


##################################################
#PARAMETERS TO NOT CHANGE
$TenantAdminURL = "https://(...)-admin.sharepoint.com/"
##################################################

$Owner = ""
$SiteURL = "https://(...).sharepoint.com/sites/qaz"

#Connect to PnP Online
Connect-PnPOnline -Url $TenantAdminURL -Interactive

Set-PnPTenantSite -Url $SiteURL -Owners $Owner
Write-Host -f Green "$Owner Adicionado ao Site" $SiteURL