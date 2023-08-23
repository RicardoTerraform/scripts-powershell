########################################################
#  https://docs.microsoft.com/en-us/microsoft-365/compliance/delete-items-in-the-recoverable-items-folder-of-mailboxes-on-hold?view=o365-worldwide#step-1-collect-information-about-the-mailbox
#  
#  STEP 1 - Collect information about the mailbox
#
#  STEP 2 - Prepare the mailbox 
#
#  STEP 3 - Remove all holds from the mailbox
#
#  STEP 4 - Remove the delay hold from the mailbox
#
#  STEP 5 - Delete items in the Recoverable Items folder
#
#  STEP 6 - Revert the mailbox to its previous state
#########################################################


# Função para aparecer uma opção de sim ou não
Function QuestionYN($param){
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription '&Yes'
    $no = New-Object System.Management.Automation.Host.ChoiceDescription '&No'
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
    $title = 'Confirm'
    $message = $param
    $result = $host.ui.PromptForChoice($title, $message, $options, 0) 
    return $result
}


if (!$ExoSession){
    Connect-ExchangeOnline}

if (!$SccSession)
   {Connect-IPPSSession}

Write-Warning "You need to get the mailbox client access settings so you can temporarily disable them so the owner (or other users) CANNOT ACCESS the mailbox during this procedure. YOU MUST WARN THE USER before you get it started."
Write-Host 
Write-Host "";"Seguir os passos no documento StepsDeleteRecoverableItems.TXT antes de continuar.";""

$message= "Os 1, 2, 3 e 4 do .TXT estão concluidos?"
$pergunta= QuestionYN $message

if($pergunta -eq 0){

    $address = Read-Host "Enter a Target Email Address"
    $searchName = "FoldersSearch"

    if ($address.IndexOf("@") -ige 0)
    {
        $folderQueries = @()
        $folderStatistics = Get-MailboxFolderStatistics $address -FolderScope RecoverableItems

        foreach ($folderStatistic in $folderStatistics)
        {
            $folderId = $folderStatistic.FolderId
            $folderPath = $folderStatistic.FolderPath
            $folderItems = $folderStatistic.ItemsInFolderAndSubfolders
            $folderSize = $folderStatistic.FolderAndSubfolderSize

            $folderStat = New-Object PSObject
            Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderPath -Value $folderPath
            Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderSize -Value $folderSize
            Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderItems -Value $folderItems

            $folderQueries += $folderStat
        }

        Write-Host "";"-------Recoverable Folders-------"
        $folderQueries |ft

        #do{$question = Read-Host "Do you want to DELETE items in the Recoverable Items folder? Enter (y / n)"
        #}while($question -ne "y" -and $question -ne "n")

        $message= "Do you want to DELETE items in the Recoverable Items folder?"
        $question= QuestionYN $message

        if($question -eq 0){
            Write-Host ""
        
        ##############################################################
        #                                                            #
        # VERIFICAR SE ALGUNS ATRIBUTOS FORAM CORRETAMENTE ALTERADOS #
        #                                                            #
        ##############################################################  


            $SingleItemRecoveryEnabled= (Get-Mailbox $address).SingleItemRecoveryEnabled
            if($SingleItemRecoveryEnabled -ne $false)
            {
                Write-Warning "Não completaste os Passos todos. O SingleItemRecoveryEnabled devia ser False e continua $SingleItemRecoveryEnabled"
                (Get-Mailbox $address).SingleItemRecoveryEnabled
                Break 
            }

            $RetainDeletedItemsFor=(Get-Mailbox $address).RetainDeletedItemsFor
            if($RetainDeletedItemsFor -ne "30.00:00:00")
            {
                Write-Warning "Não completaste os Passos todos. O RetainDeletedItemsFor devia ser 30 dias e continua com $RetainDeletedItemsFor dias"
                (Get-Mailbox $address).RetainDeletedItemsFor
                Break
            }
 
            $EwsEnabled=(Get-CASMailbox $address).EwsEnabled
            if($EwsEnabled -ne $false)
            {
                Write-Warning "Não completaste os Passos todos. O EwsEnabled devia ser False e continua com $EwsEnabled"
                (Get-Mailbox $address).EwsEnabled
                Break
            }

            $ActiveSyncEnabled=(Get-CASMailbox $address).ActiveSyncEnabled
            if($ActiveSyncEnabled -ne $false)
            {
                Write-Warning "Não completaste os Passos todos. O ActiveSyncEnabled devia ser False e continua com $ActiveSyncEnabled"
                (Get-Mailbox $address).ActiveSyncEnabled
                Break
            }

            $MAPIEnabled=(Get-CASMailbox $address).MAPIEnabled
            if($MAPIEnabled -ne $false)
            {
                Write-Warning "Não completaste os Passos todos. O MAPIEnabled devia ser False e continua com $MAPIEnabled"
                (Get-Mailbox $address).MAPIEnabled
                Break
            }

            $OWAEnabled=(Get-CASMailbox $address).OWAEnabled
            if($OWAEnabled -ne $false)
            {
                Write-Warning "Não completaste os Passos todos. O OWAEnabled devia ser False e continua com $OWAEnabled"
                (Get-Mailbox $address).OWAEnabled
                Break
            }

            $ImapEnabled=(Get-CASMailbox $address).ImapEnabled
            if($ImapEnabled -ne $false)
            {
                Write-Warning "Não completaste os Passos todos. O ImapEnabled devia ser False e continua com $ImapEnabled"
                (Get-Mailbox $address).ImapEnabled
                Break
            }

            $PopEnabled=(Get-CASMailbox $address).PopEnabled
            if($PopEnabled -ne $false)
            {
                Write-Warning "Não completaste os Passos todos. O PopEnabled devia ser False e continua com $PopEnabled"
                (Get-Mailbox $address).PopEnabled
                Break
            }

            $ElcProcessingDisabled=(Get-Mailbox $address).ElcProcessingDisabled
            if($ElcProcessingDisabled -ne $true)
            {
                Write-Warning "Não completaste os Passos todos. O ElcProcessingDisabled devia ser False e continua com $ElcProcessingDisabled"
                (Get-Mailbox $address).ElcProcessingDisabled
                Break
            }

            $LitigationHoldEnabled=(Get-Mailbox $address).LitigationHoldEnabled
            if($LitigationHoldEnabled -ne $false)
            {
                Write-Warning "Não completaste os Passos todos. O LitigationHoldEnabled devia ser False e continua com $LitigationHoldEnabled"
                (Get-Mailbox $address).LitigationHoldEnabled
                Break
            }
            
            ##############################################################
            #                                                            #
            #                          STEP 5                            #
            #                                                            #
            ##############################################################
        
            Write-Host "";"Deleting Items in the SubFolder DELETIONS , DISCOVERYHOLDS , PURGES and SUBSTRATEHOLDS";""
        
            $Recoverable = Get-MailboxFolderStatistics $address -FolderScope RecoverableItems |?{$_.FolderPath -like "*Deletions" -or $_.FolderPath -like "*DiscoveryHolds" -or $_.FolderPath -like "*Purges" -or $_.FolderPath -like "*SubstrateHolds"}
        
            :BreakAll foreach($RecoverableItems in $Recoverable){
                
                $break2=0
                $CheckItemsFolderEqual= @(0,0,0)
                $SubFolder = $RecoverableItems.Name
                $folderId = $RecoverableItems.FolderId
                $encoding= [System.Text.Encoding]::GetEncoding("us-ascii")
                $nibbler= $encoding.GetBytes("0123456789ABCDEF");
                $folderIdBytes = [Convert]::FromBase64String($folderId);
                $indexIdBytes = New-Object byte[] 48;
                $indexIdIdx=0;
                $folderIdBytes | select -skip 23 -First 24 | %{$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -shr 4];$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -band 0xF]}
                $folderQuery = "folderid:$($encoding.GetString($indexIdBytes))";
        
                Remove-ComplianceSearch $searchName -Confirm:$false -ErrorAction 'SilentlyContinue'

                Write-Host "";"Start Deleting SubFolder $SubFolder..."

                :BreakSubFolder do{
                    $count=0
                    $count_2=0
                    Write-host "";"Keep Deleting...";""
                    Remove-ComplianceSearch $searchName -Confirm:$false -ErrorAction 'SilentlyContinue'
                    $complianceSearch = New-ComplianceSearch -Name $searchName -ExchangeLocation $address -ContentMatchQuery "$folderQuery"
   
                    Start-ComplianceSearch -Identity $searchName

                    do{
                        Write-host "Waiting for search to complete..."
                        Start-Sleep -s 5
                        $count++

                        #Função para caso o "status" não se altere por mais de 1min30s, elimina essa instancia e cria um novo ciclo.
                        #Normalmente demora 30s a completar
                        if($count -gt 18)
                        {
                            Remove-ComplianceSearch $searchName -Confirm:$false -ErrorAction 'SilentlyContinue'
                            $TotalItems = 2
                            Break   
                        }
                    }while ((Get-ComplianceSearch $searchName).Status -ne 'Completed')
 
                    Write-host ""
                
                    If($count -le 18){  
                        $TotalItems=(Get-ComplianceSearch $searchName).Items

                        if ($TotalItems -gt 1)
                        {
                            $complianceSearchAction = New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType HardDelete -Confirm:$false
                            $searchActionName = ($complianceSearchAction).Name
       
                            do
                            {
                                Write-host "Waiting for search action to complete..."
                                Start-Sleep -s 5
                                $count_2++
                                
                                #Função para caso o "status" não se altere por mais de 2min30s, elimina essa instancia e cria um novo ciclo
                                #Normalmente demora 1min a completar
                                if($count_2 -gt 36)
                                {
                                    Remove-ComplianceSearch $searchName -Confirm:$false -ErrorAction 'SilentlyContinue'
                                    Break   
                                }

                            }while ((Get-ComplianceSearchAction $searchActionName).Status -ne 'Completed') 


                            
                            if($count_2 -le 36){
                                
                                $CheckItemsFolderEqual[0]=$CheckItemsFolderEqual[1]
                                $CheckItemsFolderEqual[1]=$CheckItemsFolderEqual[2]
                                $CheckItemsFolderEqual[2]=$TotalItems
                                if($CheckItemsFolderEqual[0] -eq $TotalItems -and $CheckItemsFolderEqual[1] -eq $TotalItems -and $CheckItemsFolderEqual[2] -eq $TotalItems)
                                {
                                    
                                    Write-Warning "Items da SubFolder $SubFolder NÃO estão a ser Eliminados!!! Fazer Troubleshooting..."
                                    Write-Host "";

                                    $message= "Quer continuar com a Eliminação dos Items das próximas Subfolders?"
                                    $question2= QuestionYN $message

                                    if($question2 -eq 0){
                                        
                                        Break BreakSubFolder
                                    }else{
                                        $break2=1
                                        Break BreakAll
                                    }                                 
                                }

                                #função para calcular tempo restante baseado nos items restantes
                                if($TotalItems -lt 10){
                                    Write-Host "";"SubFolder $SubFolder 0 minutes left - (0 Items to be Deleted)";"Last Double-check..."}
                                else{
                                    $aux=$TotalItems-10
                                    [int]$minutes = (((($aux) * 90)/10)/60)
                                    Write-Host "";"SubFolder $SubFolder ~$minutes minutes left - ($aux Items to be Deleted)"
                                }
                            }

                       }else{
                            Write-Host "";"No Items were found in this SubFolder $SubFolder";"______________________________________________"     
                            }
                    }
                
                }while($TotalItems -gt 1)

                Remove-ComplianceSearch $searchName -Confirm:$false -ErrorAction 'SilentlyContinue'
                #Write-Host "Items in SubFolder $SubFolder were DELETED";"______________________________________________"
                               
            }

            Get-MailboxFolderStatistics $address -FolderScope RecoverableItems |Select  Name,FolderAndSubfolderSize, ItemsInFolderAndSubfolders      
            
            if($break2 -eq 1){
                    Write-Host "";"";"The Process is NOT COMPLETED";""
                }else{
                    Write-Host "";"";"The Process is COMPLETED";"";"";"Voltar para o Documento .TXT e completar o Passo 6"
                    }
            
    
        }else{
            Write-Host "Any Recoverable Folders was Deleted"
            Break}

    }else{
        Write-Error "Couldn't recognize $address as an email address"
        }
}else{
    Write-Host "";"Abre o documento StepsDeleteRecoverableItems.txt e segue os passos todos."
    Break}