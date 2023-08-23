# Randomize passwords
function Get-RandomPassword {
    Param(
        [Parameter(mandatory = $true)]
        [int]$Length
    )
    Begin {
        if ($Length -lt 10) {
            End
        }
        $Numbers = 1..9
        $LettersLower = 'abcdefghijklmnopqrstuvwxyz'.ToCharArray()
        $LettersUpper = 'ABCEDEFHIJKLMNOPQRSTUVWXYZ'.ToCharArray()
        $Special = '!@#$%^&*()=+[{}]/?<>'.ToCharArray()

        # For the 4 character types (upper, lower, numerical, and special)
        $N_Count = [math]::Round($Length * .2)
        $L_Count = [math]::Round($Length * .4)
        $U_Count = [math]::Round($Length * .2)
        $S_Count = [math]::Round($Length * .2)
    }
    Process {
        $Pswrd = $LettersLower | Get-Random -Count $L_Count
        $Pswrd += $Numbers | Get-Random -Count $N_Count
        $Pswrd += $LettersUpper | Get-Random -Count $U_Count
        $Pswrd += $Special | Get-Random -Count $S_Count

        # If the password length isn't long enough (due to rounding), add X special characters
        # Where X is the difference between the desired length and the current length.
        if ($Pswrd.length -lt $Length) {
            $Pswrd += $Special | Get-Random -Count ($Length - $Pswrd.length)
        }

        # Lastly, grab the $Pswrd string and randomize the order
        $Pswrd = ($Pswrd | Get-Random -Count $Length) -join ""
    }
    End {
        $Pswrd
    }
}

Connect-ExchangeOnline
Connect-AzureAD

#file must be in "csv" format
#Open file/folder dialog box
Write-Host "Escolher o ficheiro Excel"
Start-Sleep -Milliseconds 90
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.initialDirectory = $initialDirectory
$OpenFileDialog.filter = "All files (*.*)| *.*"
$OpenFileDialog.ShowDialog() |  Out-Null   

$file_path = Import-Csv $OpenFileDialog.filename
#$file_path = Import-Csv C:\Users\ricardo.j.alves\Desktop\Worten\Scripts\Office\Rooms.csv

$total=0
$i=0
foreach($f in $file_path)
{$total++}

#variavel para export
$ListItemCollection = @() 

foreach($file in $file_path)
{
    #Write-Progress#
    $i++
    [int]$Completed = ($i/$total*100)
    Write-Progress -Activity "Create MailBox-Room in Progress ($i out of $total)..." -Status "$Completed% Complete:" -PercentComplete $Completed
    Start-Sleep -Milliseconds 250
    
    #Get information from CSV#
    $videoC=$file.VC
    $name = $file.Alias
    $Displayname = $file.DisplayName
    $email = $file.Email
    $capacity = $file.Capacity
    $location = $file.Location
    #$BookingWindowInDays = $file.
    #$MaximumDurationInMinutes = $file.


    #Password Generator, tamanaho 16
    #$password = Get-RandomPassword -Length 16
    #$password

 
    #Create a new "Room" mailbox#

    <#if($videoC -eq "YES"){
        #Password Generator, tamanaho 16
        $password = Get-RandomPassword -Length 16
        New-Mailbox -DisplayName $Displayname -Name $name -PrimarySmtpAddress $email -Password (ConvertTo-SecureString -String $password -AsPlainText -Force) -ResetPasswordOnNextLogon $false -Room
    }else{
        New-Mailbox -DisplayName $Displayname -Name $name -PrimarySmtpAddress $email -Room
        $password = $null
    }#>
    $password = "WRT#vc2022"
    New-Mailbox -DisplayName $Displayname -Name $name -PrimarySmtpAddress $email -Password (ConvertTo-SecureString -String $password -AsPlainText -Force) -ResetPasswordOnNextLogon $false -Room
    Start-Sleep -Seconds 10
    #Set-CalendarProcessing -AllowRecurringMeetings $True/$False -ScheduleOnlyDuringWorkHours $True/$False -BookingWindowInDays $BookingWindowInDays -MaximumDurationInMinutes $MaximumDurationInMinutes
    #Set-CalendarProcessing -Identity $email -ResourceDelegates "rjalves@ext.worten.pt"
    Set-CalendarProcessing -Identity $email -DeleteComments $false -DeleteSubject $false -AddOrganizerToSubject $false -ProcessExternalMeetingMessages $True
    Set-Place -Identity $email -Capacity $capacity
    Set-Mailbox -Identity $email -Office $location


    Write-Host " $email's MailBox Created..."

    Start-Sleep -Seconds 3 
    

    $ExportItem = New-Object PSObject
    $ExportItem | Add-Member -MemberType NoteProperty -name "DisplayName" -value $Displayname
    $ExportItem | Add-Member -MemberType NoteProperty -Name "Email" -value $email
    $ExportItem | Add-Member -MemberType NoteProperty -Name "Password" -value $password
   
    #Add the object with the above properties to the Array
    $ListItemCollection += $ExportItem
     
}
#create a export file
$ListItemCollection | Export-csv -Path "C:\Users\ricardo.j.alves\Desktop\GroupPasswords.csv" -NoTypeInformation 

Write-Host " Waiting 2min30s to change the UPN" 
Start-Sleep -Seconds 180



#Change UPN Azure AD#
foreach($files in $file_path)
{
    $emails = $files.Email
    #Get-AzureADUser -SearchString $emails
    $localuser = Get-AzureADUser -SearchString $emails
    $localuser | foreach {$newUpn = $_.UserPrincipalName.Replace("worten.onmicrosoft.com","worten.pt"); $_ | Set-AzureADUser -UserPrincipalName $newUpn}
    #Get-AzureADUser -SearchString $emails
}

Start-Sleep -Seconds 30

#adicionar a um grupo de Licensing-Microsoft-TeamsRooms
foreach($filess in $file_path)
{
    $vc=$filess.VC
    if($vc -eq "YES"){
        $emails = $filess.Email
        $obID=Get-AzureADUser -SearchString $emails  | Select-Object -ExpandProperty Objectid
        Add-AzureADGroupMember -ObjectId ada497fb-15e9-47c8-be9b-5c1a26a23b0e -RefObjectId $obID
        Write-Host " $emails's MailBox ADDED to the License-Group  Licensing-Microsoft-TeamsRooms..."
    }
}


Write-Host "";" DONE" 

