#############################################################################################################
#                                                                                                           #
#    Script vai correr todos os Devices Android                                                             #
#    Vai eliminar os devices com o mesmo Serial Number mantendo o device com o Last sync date mais recente  #
#                                                                                                           #
#############################################################################################################


Connect-MSGraph

#Connection Azure AD
try
{
    $TestAzureADConnection = Get-AzureADCurrentSessionInfo -ErrorAction Stop
}
catch
{
    Connect-AzureAD 
}


#vai buscar todos os android
$AllDevices = Get-IntuneManagedDevice | Get-MSGraphAllPages | Where {$_.deviceName -like "*_AndroidEnterprise_*"}
$count=0
$save = New-Object System.Collections.Generic.List[System.Object]
foreach ($Devices in $AllDevices)
{

    $DeviceSerialNumber=$Devices.serialNumber
    $i=0
  
    #Validar SN Repetidos
    if ( $save.contains($DeviceSerialNumber))
    {
        # Serial Number JÁ EXISTE no array "$save", significa que não precisa de repetir o processo, já foram eliminados os devices deste Serial number
        Write-Host "Já foram eliminados os Devices deste SN "$DeviceSerialNumber
    }
    else
    {
        # Serial Number NÃO EXISTE no array "$save", significa que precisa de correr o processo.
        
        $CountDevice=(Get-IntuneManagedDevice -Filter "serialNumber eq '$DeviceSerialNumber'").count


        if($CountDevice -gt 1)
        {
            Write-Host "$count) Device $DeviceSerialNumber -" ($CountDevice-1) " devices serão removidos" -f Yellow
       
            $Device= Get-IntuneManagedDevice -Filter "serialNumber eq '$DeviceSerialNumber'"
            $ordenar=$Device | Sort-Object -Property lastSyncDateTime
 
            For($i=0; $i -lt ($CountDevice-1); $i++)
            {
                $AzureADDeviceDeviceID = $ordenar[$i].id

                $AzureADDeviceObjectID = (Get-AzureADDevice -All:$true | Where-Object {$_.DeviceId -eq $AzureADDeviceDeviceID}).ObjectId

                #remover do azure
                Remove-AzureADDevice -ObjectId $AzureADDeviceObjectID
                Write-Host "Removido do Azure" -f Green
                
                #remover do Intune
                Remove-IntuneManagedDevice –managedDeviceId $AzureADDeviceDeviceID
                Write-Host "Removido do Intune" -f Green


            }
       }

       $count ++          
    }

    #É adicionado ao Array $save o SN para comparação
    $save.Add($DeviceSerialNumber)
} 
