Connect-ExchangeOnline

$Room = get-Mailbox -resultsize unlimited | where { $_.RoomMailboxAccountEnabled -match "True"}

foreach($rooms in $Room)
{
    if($rooms.DisplayName -like "WTC*")
    {
        Remove-MailBox -identity $rooms.DisplayName
    }
}