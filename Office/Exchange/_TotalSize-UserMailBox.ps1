$Users = Get-Mailbox -ResultSize unlimited 
$result = @()
$totalmbx = $Users.Count
$i = 0
Foreach ($user in $Users)
{
    $i++
    Write-Progress -activity "Processing $user.DisplayName" -status "$i out of $totalmbx completed"
    $stat= Get-MailboxStatistics -Identity $user.PrimarySMTPAddress 
    $size =[math]::Round(($stat.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1MB),2)
    if ($size -gt 70000)
    {
    
        #$result = New-Object PSObject
        #Add-Member -InputObject $result -MemberType NoteProperty -Name Name -Value $user.DisplayName
        #Add-Member -InputObject $result -MemberType NoteProperty -Name FolderSize -Value $size
        #$result
        Out-File -FilePath "C:\temp\TotalSize.csv" -InputObject "$($user.PrimarySMTPAddress), $size" -Encoding UTF8 -append
    }
}  

