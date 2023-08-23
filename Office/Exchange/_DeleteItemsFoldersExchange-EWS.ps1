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

do{
    $address = Read-Host "Enter a Target Email Address"
    $count =  @(Get-MailboxFolderStatistics -Identity $address).count

}while($count -lt 2)


$List =  @(Get-MailboxFolderStatistics -Identity $address)

[System.Console]::CursorVisible = $false 

#menu offset to allow space to write a message above the menu
$xmin = 3
$ymin = 5
 
#Write Menu
Clear-Host
Write-Host "    $address's MailBox" -ForegroundColor Green
Write-Host ""
Write-Host "    Use the up / down arrow to navigate and Enter to make a selection. 'Enter' - Run, 'ESC' - Exit`n"
[Console]::SetCursorPosition(0, $ymin)


$x=0
        $x1=0
        foreach ($names in $List){
            $valor= $names.FolderPath.Length
            $valor1= $names.FolderAndSubfolderSize.Length

            if($valor -gt $x)
                {
                    $x = $valor
                } 
                
             if($valor1 -gt $x1)
                {
                    $x1 = $valor1
                }    
        }
        $coordenadaX1=$x+13
        $coordenadaX2=$coordenadaX1+$x1+7


        [console]::setcursorposition(7,3),[System.Console]::WriteLine("Folder Path") 
        [console]::setcursorposition($coordenadaX1,3), [System.Console]::WriteLine("Folder Size")
        [console]::setcursorposition($coordenadaX2,3), [System.Console]::WriteLine("Folder Items")
        [console]::setcursorposition(6,4), [System.Console]::WriteLine("-------------------------------------------------------------------------------------------------------")

$ymin2=5
foreach ($name in $List) {
    for ($i = 0; $i -lt $xmin; $i++) {
        Write-Host " " -NoNewline
    }
    Write-Host "   " + $name.FolderPath -NoNewline
    [console]::setcursorposition($coordenadaX1,$ymin2) 
    Write-Host $name.FolderAndSubfolderSize
    [console]::setcursorposition($coordenadaX2,$ymin2) 
    Write-Host $name.ItemsInFolderAndSubfolders

    $ymin2++
}
 
#Highlights the selected line
function Write-Highlighted {

    [Console]::SetCursorPosition(0,0)
    [Console]::SetCursorPosition(1 + $xmin, $cursorY + $ymin)
    
    Write-Host ">" -BackgroundColor White -ForegroundColor DarkMagenta -NoNewline
    Write-Host " " + $List[$cursorY].FolderPath -BackgroundColor White -ForegroundColor DarkMagenta -NoNewline
    [console]::setcursorposition($coordenadaX1,$cursorY +$ymin) 
    Write-Host $List[$cursorY].FolderAndSubfolderSize -BackgroundColor White -ForegroundColor DarkMagenta -NoNewline
    [console]::setcursorposition($coordenadaX2,$cursorY +$ymin) 
    Write-Host $List[$cursorY].ItemsInFolderAndSubfolders -BackgroundColor White -ForegroundColor DarkMagenta
    [console]::setcursorposition($coordenadaX2,$cursorY +$ymin) 
    [Console]::SetCursorPosition(0, $cursorY + $ymin)  
       
}
 
#Undoes highlight
function Write-Normal {
    [Console]::SetCursorPosition(1 + $xmin, $cursorY + $ymin)
    Write-Host "  " + $List[$cursorY].FolderPath -NoNewline
    [console]::setcursorposition($coordenadaX1,$cursorY +$ymin)
    Write-Host $List[$cursorY].FolderAndSubfolderSize -NoNewline
    [console]::setcursorposition($coordenadaX2,$cursorY +$ymin)
    Write-Host $List[$cursorY].ItemsInFolderAndSubfolders 
}
   
#highlight first item by default
$cursorY = 0
Write-Highlighted

$selection = ""
    
:Breakall do{    
 
    if ([console]::KeyAvailable) {
        $x = $Host.UI.RawUI.ReadKey()
        [Console]::SetCursorPosition(1, $cursorY)
        Write-Normal
        switch ($x.VirtualKeyCode) { 
            38 {
                #down key
                if ($cursorY -gt 0) {
                    $cursorY = $cursorY - 1
                }
            }
 
            40 {
                #up key
                if ($cursorY -lt $List.Length - 1) {
                    $cursorY = $cursorY + 1
                }
            }

            13 {
                #enter key
                Write-Highlighted
                $selection = $List[$cursorY].FolderPath
                [Console]::SetCursorPosition(3, $ymin2)
                $message = "Tem a certeza que pretende ELIMINAR os items da pasta $selection ?"
                $pergunta = QuestionYN $message

                if($pergunta -eq 1){
                    Write-Host "";"Processo interrompido..."
                    $pergunta=10
                    BREAK Breakall
                }
                else{ 
                    BREAK Breakall}
            }

            27 {
                #ESC key
                [Console]::SetCursorPosition(3, $ymin2)
                Write-Host "";"Processo interrompido..."
                $pergunta=10
                BREAK Breakall
            }
        }
        Write-Highlighted
    }
    Start-Sleep -Milliseconds 5 #Prevents CPU usage from spiking while looping
 }while (([System.Int16]$inputChar.Key -ne [System.ConsoleKey]::Enter) -and ([System.Int16]$inputChar.Key -ne [System.ConsoleKey]::Escape))


$result = $selection

#Clear-Host

if($pergunta -eq 0){
    
    $dllpath = “C:\Program Files\Microsoft Office\Office16\ADDINS\Microsoft Power Query for Excel Integrated\bin\Microsoft.Exchange.WebServices.dll”
    [VOID][Reflection.Assembly]::LoadFile($dllpath)


    $credentials=Get-Credential -Message "Please enter your admin credentials"
    Write-Host "Connecting ... (It might take up 2min)"
    $Username = $credentials.username
    $Password = $credentials.GetNetworkCredential().Password | ConvertTo-SecureString -AsPlainText -Force
    $secureStringText = $Password | ConvertFrom-SecureString
    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016)
    $Service.Credentials = New-Object System.Net.NetworkCredential ($Username,$Password)

    $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $address)
    $Service.AutodiscoverUrl($address, {$true})

    $Service.Url = "https://outlook.office365.com/EWS/Exchange.asmx"


    Write-Host "Processing Mailbox: $address" -ForegroundColor Green
    
    Write-Host "";"Before Delete"
    Get-MailboxFolderStatistics -Identity $address | ? {$_.FolderPath -eq $result}| Select FolderPath,FolderAndSubfolderSize, ItemsInFolderAndSubfolders

    get-mailboxfolderstatistics $address | Where-Object{$_.FolderPath -eq $result } | ForEach-Object{
        Add-Type -AssemblyName System.Web
        $urlEncodedId = [System.Web.HttpUtility]::UrlEncode($_.FolderId.ToString())


        $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId((Convertid $urlEncodedId $address))
        $ewsFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
        }

    #$ewsFolder

    $pause=0
    do{
        $pause++
        Start-Sleep -Seconds 5
        Write-Host "Eliminando..."
    }While($pause -le 3)

    $ewsFolder.Empty([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete, $false)
    #$ewsFolder.Empty([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete, $false)

    $pause1=0
    do{
        $pause1++
        Start-Sleep -Seconds 5
        Write-Host "Eliminando..."
    }While($pause1 -le 5)

    Write-Host "";"After Delete"
    Get-MailboxFolderStatistics -Identity $address | ? {$_.FolderPath -eq $result}| Select FolderPath,FolderAndSubfolderSize, ItemsInFolderAndSubfolders


    Write-Host "";"Os Items da Pasta $result foram todos eliminados"


}else{
    Break
}
