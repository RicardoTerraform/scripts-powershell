#######################################################################################################################
#
#  Script vai alterar o "Version History" da Libraries mais importantes de todos os Sites.
#  Está por Default que o nº máximo de versão é 5, a variável pode ser alterada "$VersioningLimit".
#  O user Admin que executar este código tem de estar adicionado à Enterprice APP "PnP Management Shell" - App ID: 31359c7f-bd7e-475c-86db-fdb8c937548e e fazer um user consent"
#  Só admins podem ser membros desta App "PnP Management Shell"
#
#######################################################################################################################


##################################################
#PARAMETERS TO NOT CHANGE
#Set Variables
$TenantAdminURL = "https://(...)-admin.sharepoint.com/"

$Exclude = "POINTPUBLISHINGTOPIC#0","POINTPUBLISHINGHUB#0", "TEAMCHANNEL#1", "TEAMCHANNEL#0", "RedirectSite#0","SPSMSITEHOST#0", "EDISC#0","STS#-1","SRCHCEN#0"

$countSite=1
###################################################

$VersioningLimit=5

#Connect
Connect-PnPOnline -Url $TenantAdminURL -Interactive

$SitesURL = Get-PnPTenantSite | ? {$_.Template -notin $Exclude}

ForEach($Site in $SitesURL)
{
    Try {
        Write-host -f Yellow "$countSite - Site "$Site.URL" has been Updated"
        Connect-PnPOnline -Url $Site.URL -Interactive
 
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
                Write-host -f Green "`tVersion History Settings has been Updated on '$($Library.Title)'"
            }
            Else
            {
                Write-host -f Yellow "`tVersion History is turned-off at '$($Library.Title)'"
            }
        }
    }
    Catch {
        Write-host -f Red "Não foi possivel Alterar o número de versões para este site "$Site.URL
        Write-host -f Red "Error:" $_.Exception.Message     
    }

    $countSite += 1
}
