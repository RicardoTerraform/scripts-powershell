#Import-Module ExchangeOnlineManagement
if (!$ExoSession){
Connect-ExchangeOnline}

if (!$SccSession)
   {Connect-IPPSSession}

$address = Read-Host "Enter an email address"
$searchName = "FoldersSearch"

if ($address.IndexOf("@") -ige 0)
{
    $folderQueries = @()
    $folderStatistics = Get-MailboxFolderStatistics $address

    foreach ($folderStatistic in $folderStatistics)
    {
        $folderId = $folderStatistic.FolderId
        $folderPath = $folderStatistic.FolderPath
        $folderItems = $folderStatistic.ItemsInFolder
        $folderSize = $folderStatistic.FolderSize

        $folderStat = New-Object PSObject
        Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderPath -Value $folderPath
        #Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderQuery -Value $folderQuery
        Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderSize -Value $folderSize
        Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderItems -Value $folderItems

        $folderQueries += $folderStat
    }

    Write-Host "-----Exchange Folders-----"
    $folderQueries |ft

    $question = Read-Host "Do you want to DELETE any Folder? Enter (yes / no)"
    
    if($question -eq "yes"){
        do{
            $folderchoose = Read-Host "Enter the FULL FOLDER PATH that you want to Delete as you can see on the screen"

            foreach($folderStatistic in $folderStatistics)
            {
                if($folderchoose -eq $folderStatistic.FolderPath)
                {
                    $aux=1

                    $folderId = $folderStatistic.FolderId

                    $encoding= [System.Text.Encoding]::GetEncoding("us-ascii")
                    $nibbler= $encoding.GetBytes("0123456789ABCDEF");
                    $folderIdBytes = [Convert]::FromBase64String($folderId);
                    $indexIdBytes = New-Object byte[] 48;
                    $indexIdIdx=0;

                    $folderIdBytes | select -skip 23 -First 24 | %{$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -shr 4];$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -band 0xF]}

                    $folderQuery = "folderid:$($encoding.GetString($indexIdBytes))";

                    break
                }else{$aux=2}
           }   
        }while ($aux -eq 2)


    Remove-ComplianceSearch $searchName -Confirm:$false -ErrorAction 'SilentlyContinue'

    do{
        Write-host ""
        Remove-ComplianceSearch $searchName -Confirm:$false -ErrorAction 'SilentlyContinue'
        $complianceSearch = New-ComplianceSearch -Name $searchName -ExchangeLocation $address -ContentMatchQuery "$folderQuery"
   
        Start-ComplianceSearch -Identity $searchName

        do{
            Write-host "Waiting for search to complete..."
            Start-Sleep -s 5
            #$complianceSearch = Get-ComplianceSearch $searchName
        }while ((Get-ComplianceSearch $searchName).Status -ne 'Completed')
 
        Write-host ""  
   
        if ((Get-ComplianceSearch $searchName).Items -gt 0)
        {
            $complianceSearchAction=New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType SoftDelete -Confirm:$false
            $searchActionName=($complianceSearchAction).Name
       
            do
            {
                Write-host "Waiting for search action to complete..."
                Start-Sleep -s 5
                #$complianceSearchAction = Get-ComplianceSearchAction $searchActionName
            }while ((Get-ComplianceSearchAction $searchActionName).Status -ne 'Completed') 

            (Get-ComplianceSearchAction $searchActionName).Results
            #Remove-ComplianceSearch $searchName -Confirm:$false -ErrorAction 'SilentlyContinue'

            $items = (Get-MailboxFolderStatistics -Identity $address |?{$_.FolderId -eq $folderId}).ItemsInFolder
            [int]$minutes = ((($items * 90)/10)/60)
            Write-Host "";"About $minutes minutes left"

        }else{

                Write-Host "";"No Items were found in this folder $folderchoose";""
                Get-MailboxFolderStatistics -Identity $address |Where {$_.FolderId -eq $folderId}  | Select Name,FolderPath,FolderSize, ItemsinFolder
                Write-Host "The Process is COMPLETED"
        }
    }while((Get-ComplianceSearch $searchName).Items -gt 0)


    Remove-ComplianceSearch $searchName -Confirm:$false -ErrorAction 'SilentlyContinue'

    }else{Break}

}else{
   Write-Error "Couldn't recognize $address as an email address"
}

#Disconnect-ExchangeOnline
#Disconnect-PSSession
