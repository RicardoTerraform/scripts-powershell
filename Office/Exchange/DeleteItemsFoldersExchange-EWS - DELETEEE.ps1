function ConvertId{
	param (
	        $OwaId = "$( throw 'OWAId is a mandatory Parameter' )",
            $Mailbox
		  )
	process{
	    $aiItem = New-Object Microsoft.Exchange.WebServices.Data.AlternateId
	    $aiItem.Mailbox = $Mailbox
	    $aiItem.UniqueId = $OwaId
	    $aiItem.Format = [Microsoft.Exchange.WebServices.Data.IdFormat]::OwaId
	    $convertedId = $service.ConvertId($aiItem, [Microsoft.Exchange.WebServices.Data.IdFormat]::EwsId)
		return $convertedId.UniqueId
	}
}



Connect-ExchangeOnline
#TargetMail
$MailboxNames = ""
#TargetFolder
#$FolderDelete = "Itens Eliminados"
#AdminAccount
$AdminID= ""
#Enter Admin Password
$AdminPwd = Read-Host "Enter Password" -AsSecureString


$dllpath = “C:\Program Files\Microsoft Office\Office16\ADDINS\Microsoft Power Query for Excel Integrated\bin\Microsoft.Exchange.WebServices.dll”
[VOID][Reflection.Assembly]::LoadFile($dllpath)

$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016)
$Service.Credentials = New-Object System.Net.NetworkCredential ($AdminID, $AdminPwd)

$Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxNames)
$service.AutodiscoverUrl($MailboxNames, {$true})

$Service.Url = "https://outlook.office365.com/EWS/Exchange.asmx"


Write-Host "Processing Mailbox: $MailboxNames" -ForegroundColor Green

#$RootFolderID = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $MailboxNames)
#$RootFolderID.ChangeKey


#$RootFolder= [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $RootFolderID)
#$RootFolder 






get-mailboxfolderstatistics $MailboxNames | Where-Object{$_.FolderPath -eq "/Drafts/_Alerts_Azure" } | ForEach-Object{
Add-Type -AssemblyName System.Web
$urlEncodedId = [System.Web.HttpUtility]::UrlEncode($_.FolderId.ToString())


$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId((Convertid $urlEncodedId $MailboxNames))
$ewsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
}

$ewsFolder


$FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
#$FolderView | fl
$FolderView.Traversal =  [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
$FolderView.Traversal


#$response = $ewsFolder.FindFolders($FolderView)
#$response




#$resultado = Get-MailboxFolderStatistics -Identity $MailboxNames | Where-Object {$_.FolderPath -eq "/Inbox/_Alerts_Azure"} | Select-Object -ExpandProperty Name
#[string]$resultado



#$Folder=$response | ? {$_.DisplayName -eq $resultado}
#$Folder | Select DisplayName, TotalCount
#Write-Host "Before Delete"
#Get-MailboxFolderStatistics -Identity $MailboxNames | ? {$_.Name -eq $FolderDelete}| Select Name,FolderSize, ItemsinFolder

$ewsFolder.Empty([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete, $false)
#$Folder.Empty([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete, $false)

#$Folder.Load()
#Write-Host "After Delete"
#Get-MailboxFolderStatistics -Identity $MailboxNames | ? {$_.Name -eq $FolderDelete}| Select Name,FolderSize, ItemsinFolder