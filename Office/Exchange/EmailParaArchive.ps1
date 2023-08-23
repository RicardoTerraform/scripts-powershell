Function ValidarV2 ($ItemCountv2, $TotalItemSizev2, $User)
{
    Start-Sleep -Seconds 300
    $ItemCountv3= (Get-MailboxStatistics -identity $User -Archive).ItemCount
    $TotalItemSizev3= (Get-MailboxStatistics -identity $User -Archive).TotalItemSize

    if($ItemCountv2 -ne $ItemCountv3 -and $TotalItemSizev2 -ne $TotalItemSizev3)
    {
       $result2 = 0
       Return $result2
    }
    else
    {
        Start-ManagedFolderAssistant $User
        Start-Sleep -Seconds 300

        $ItemCountv4= (Get-MailboxStatistics -identity $User -Archive).ItemCount
        $TotalItemSizev4= (Get-MailboxStatistics -identity $User -Archive).TotalItemSize

        if($ItemCountv3 -ne $ItemCountv4 -and $TotalItemSizev3 -ne $TotalItemSizev4)
        {
            $result2 = 0
            Return $result2
        }
        else
        {
            $result2 = 1
            Return $result2
        }
    }
}

Function Validar ($ItemCount, $TotalItemSize, $User)
{
    #Esperar 1h
    Start-Sleep -Seconds 3600

    $ItemCountv2= (Get-MailboxStatistics -identity $User -Archive).ItemCount
    $TotalItemSizev2= (Get-MailboxStatistics -identity $User -Archive).TotalItemSize
    
    if($ItemCount -ne $ItemCountv2 -and $TotalItemSize -ne $TotalItemSizev2)
    {
       $result = 0
       Return $result
    }
    else
    {
        Start-ManagedFolderAssistant $User
        $result1 = ValidarV2 $ItemCountv2 $TotalItemSizev2 $User
        return $result1
    }
}

################################################################################################
################################################################################################

$User=""

$ItemCount= (Get-MailboxStatistics -identity $User -Archive).ItemCount
$TotalItemSize= (Get-MailboxStatistics -identity $User -Archive).TotalItemSize.Value

Set-Mailbox -Identity $User -RetentionPolicy "One Tag Policy"

#Validar que a politica de retenção foi alterada
do{
    #Esperar 5 minutos para mudar a Retention Policy
    Start-Sleep -Seconds 300

    $validarRetention = (Get-Mailbox -Identity $User).RetentionPolicy

}while ($validarRetention -notlike "One Tag Policy")


#Validar 
do{
    Start-ManagedFolderAssistant $User
    $ResultadoFinal = Validar $ItemCount $TotalItemSize $User

}while($ResultadoFinal = 0)


Write-Host "Ficheiros movidos para o Archive está completo"


Set-Mailbox -Identity $User -RetentionPolicy "Default MRM Policy"
#Validar que a politica de retenção foi novamente alterada
do{
    #Esperar 5 minutos para mudar a Retention Policy
    Start-Sleep -Seconds 300

    $validarRetention = (Get-Mailbox -Identity $User).RetentionPolicy

}while ($validarRetention -notlike "Default MRM Policy")

