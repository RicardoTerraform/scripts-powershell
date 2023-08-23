###########################################################################################################
#
# Script vai percorrer apenas um site.
# Verifica o número de versões e elimina tudo que seja mais doque o número definido na variável $VersionsToKeep
# Também valida se o ficheiro é Major Version ou Minor Version e procede à sua eliminação
# Especificar qual o número de versões para manter na variável $VersionsToKeep
# Especificar qual o site $SitesURL
#
# Validar que o user admin que executar este scrip DEVE SER "Site Admin" de todos os sites
# O user Admin que executar este código tem de estar adicionado à Enterprice APP "PnP Management Shell" - App ID: 31359c7f-bd7e-475c-86db-fdb8c937548e e fazer um user consent"
# Só admins podem ser membros desta App "PnP Management Shell"
#
############################################################################################################

#Get Stored Credentials
#$CrescentCred = Get-AutomationPSCredential -Name "StorageRptCred"

################################################################################################


################################################################################################
#PARAMETERS TO NOT CHANGE
$MinorVersionsCount=0
$MinorVersionsToDeleteGTCount=0
$minorVersionsToDeleteLTCount = 0
$MinorVersionsToDelete = 0
$VersionsToDelete = 0
$VersionsCount=0
$minorTotal=0
$countSite=1
$save = @()
$TenantAdminURL = "https://worten-admin.sharepoint.com/"

#Exclude files type
$SystemItem ='\.(PDF|PNG|JPG|JPEG|MSG|EXE|HTML)$'

#Array to exclude system libraries
$SystemLibraries = @("Form Templates", "Pages", "Preservation Hold Library","Site Assets", "Site Pages", "Images","Site Collection Documents", "Site Collection Images","Style Library","Recursos do Site","Páginas do Site")

$Exclude = "POINTPUBLISHINGTOPIC#0","POINTPUBLISHINGHUB#0", "TEAMCHANNEL#1", "TEAMCHANNEL#0", "RedirectSite#0","SPSMSITEHOST#0", "EDISC#0","STS#-1","SRCHCEN#0"

$Query = "<View Scope='Recursive'> 
	<Query> 
		<OrderBy> 
			<FieldRef Name='ID' Ascending='TRUE' /> 
		</OrderBy>
	</Query> 
<ViewFields> <FieldRef Name='Id' /> 
</ViewFields>
<RowLimit>5000</RowLimit>
</View>"

################################################################################################

#SharePoint does not permit to delete the last versions Histories.
#If $VersionsToKeep=10, SharePoint will keep a total of 11 versions
$VersionsToKeep = 5
$SitesURL = "https://worten.sharepoint.com/sites/EquipaMarketplaceWorten-DigitalBusiness"


Connect-PnPOnline -Url $TenantAdminURL -Interactive

ForEach($Site in $SitesURL)
{
    $count=0
    Write-host -f Yellow "Site "$Site" has been Updated"
    
    Connect-PnPOnline -Url $Site -Interactive
  
    #Get the Context
    $Ctx= Get-PnPContext
    
    $DocumentLibraries = Get-PnPList | Where-Object {$_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $False -and $_.Title -notin $SystemLibraries}
   
    ForEach($Library in $DocumentLibraries)
    {
        
            $ListItems = Get-PnPListItem -List $Library -Query $Query

            $GetItem =  $ListItems | Where {$_["FileLeafRef"] -notmatch $SystemItem -and $_.FileSystemObjectType -eq "File"}

            foreach($Item in $GetItem)
            {
                $count = $count +1
                #Get File Versions
                $File = $Item.File
                $Versions = $File.Versions
                $Ctx.Load($File)
                $Ctx.Load($Versions)

                try{
                    $Ctx.ExecuteQuery()

                    Write-host -f Yellow "$count-Scanning File:"$File.Name

                    ##################################################################################
                                                    #MAJOR VERSION
                    #Get all major versions
                    #Note, the version ID goes up by 512 for each major version.
                    $majorVersions = $Versions | ? { $_.Id % 512 -eq 0}
 
                    #get the largest version number
                    $latestMajorID = $majorVersions | select Id -ExpandProperty Id | sort -Descending | select -First 1
    
                    $majorVersionsToDelete = $majorVersions | ? {$_.Id -le ($latestMajorID - 512 * $numberOfMajorVersions)}
                    #$majorVersionsToDelete

                    $VersionsCount = $majorVersionsToDelete.Count
                    $VersionsToDelete = $VersionsCount - $VersionsToKeep


                    If($VersionsToDelete -gt 0)
                    {
                        write-host -f CYAN "`t Total Number of Major Versions of the File:" $VersionsCount
                        #Delete versions
                        For($i=0; $i -lt $VersionsToDelete; $i++)
                        {
                            write-host -f Red "`t Deleting Major Version:" $majorVersions[$i].VersionLabel
                            $majorVersions[$i].DeleteObject()
                        }

                        $Ctx.ExecuteQuery()
                        #Write-Host -f Green "`t $VersionsToDelete Major Version(s) are cleaned from The FILE's Version History:"$File.Name
                    }else
                    {
                        $VersionsToDelete=0
                    }

                    ##################################################################################

                                                    #MINOR VERSION 

    
                    #Get the whole Minor Versions kept
                    $minorVersions = $Versions | ? { $_.Id % 512 -ne 0}

                    #Count the whole Minor versions
                    $MinorVersionsCount = $minorVersions.Count 
    

                    #Get Minor versions greater than the last major version kept
                    $MinorVersionsToDeleteGT = $minorVersions | ? {$_.Id -gt $latestMajorID}
                    $MinorVersionsToDeleteGTCount = $MinorVersionsToDeleteGT.Count

                    $MinorVersionsToDelete = $MinorVersionsToDeleteGTCount - $VersionsToKeep


                    #Delete Minor versions Less than the last major version kept
                    $minorVersionsToDeleteLT = $minorVersions | ? {$_.Id -lt $latestMajorID}
                    $minorVersionsToDeleteLTCount = $minorVersionsToDeleteLT.Count


                    If($MinorVersionsCount -gt 0)
                    {
                        write-host -f Cyan "`t Total Number of Minor Versions of the File:" $MinorVersionsCount
        
                        #It will delete the whole minor versions Less than the last major version kept
                        if($minorVersionsToDeleteLTCount -gt 0)
                        {
                            #Delete Minor versions
                            For($i=0; $i -lt $minorVersionsToDeleteLTCount; $i++)
                            {
                                write-host -f Red "`t Deleting Minor Version:" $minorVersionsToDeleteLT[$i].VersionLabel
                                $minorVersionsToDeleteLT[$i].DeleteObject()
                            }
                        }
                        else
                        {
                            $minorVersionsToDeleteLTCount = 0
                        }

                        $Ctx.ExecuteQuery()


                        #It will delete only the minor version higher than $VersionsToKeep = 5 on the current Major version
                        if($MinorVersionsToDelete -gt 0)
                        {
                            #Delete Minor versions
                            For($i=0; $i -lt $MinorVersionsToDelete; $i++)
                            {
                                write-host -f Red "`t Deleting Minor Version:" $MinorVersionsToDeleteGT[$i].VersionLabel
                                $MinorVersionsToDeleteGT[$i].DeleteObject()
                            }
                        }
                        else
                        {
                            $MinorVersionsToDelete = 0
                        }

                        $Ctx.ExecuteQuery()

                        #Count Minor versions were cleaned
                        $minorTotal = $minorVersionsToDeleteLTCount + $MinorVersionsToDelete

                        Write-Host -f Green "`t $VersionsToDelete Major Version(s) are cleaned from The FILE's Version History:"$File.Name
                        Write-Host -f Green "`t $minorTotal Minor Version(s) are cleaned from The FILE's Version History:"$File.Name
                    }
                    else{
                        $minorTotal=0

                        $CountMinorFinal= $MinorVersionsCountLOWER + $MinorVersionsCountGREATER
                        Write-Host -f Green "`t $VersionsToDelete Major Version(s) are cleaned from The FILE's Version History:"$File.Name
                        Write-Host -f Green "`t $minorTotal Minor Version(s) are cleaned from The FILE's Version History:"$File.Name
                        }

     
                    }catch{
                        #caso o tamanho das versões seja maior que 2GB, o script elimina as versões todas
                        $File.Versions.DeleteAll()
                        try {
                            Invoke-PnPQuery
                            Write-Host -f Green "Version history DELETED"
                            }
                        catch {
                            Write-Host -f Red " ERROR! Não foram eliminadas"
                            $File = $Item.File
                            $Ctx.Load($File)
                            $Ctx.ExecuteQuery()
                            
                            $save += [PSCUSTOMOBJECT] @{
                                    "Site" = $Site;
                                    "Documento" = $File;
                                    }
                            }



                        }
                }     
    }
}

$save