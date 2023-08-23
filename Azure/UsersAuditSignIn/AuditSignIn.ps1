Connect-AzureAD

#import data from excel
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
  InitialDirectory = [Environment]::GetFolderPath('Desktop')
  Filter           = 'Excel (*.CSV)|*.CSV'
}
$Null = $FileBrowser.ShowDialog()

$DATA= Import-Csv -Path $FileBrowser.FileName -Header 1


#$Data = "rjalves@ext.worten.pt"
$save = @()

foreach($user in $DATA.1)
{
    $GetAllLogs = Get-AzureADAuditSignInLogs -Filter "userPrincipalName eq '$user'"
    
    foreach($GetLogs in $GetAllLogs)
    {
        $Location = $GetLogs.Location.CountryOrRegion + "/" + $GetLogs.Location.State
        $date = $GetLogs.CreatedDateTime
        $application = $GetLogs.AppDisplayName
        $ip = $GetLogs.IpAddress
        $resource = $GetLogs.ResourceDisplayName
        $status = $GetLogs.ConditionalAccessStatus

        $save += [PSCUSTOMOBJECT] @{
                "UPN" = $user;
                "Date" = $date;
                "Application" = $application;
                "Status" = $status;
                "IP"=$ip
                "Country" = $Location
                "Resource" = $resource
                }
    }
}

#exportar as imformações para um CSV
$save | Export-csv -Path "C:\temp\UsersAuditSignIn.csv" -NoTypeInformation
