Connect-ExchangeOnline

$file_path= Get-Content C:\Users\ricardo.j.alves\Desktop\dados.txt

foreach($file in $file_path)
{
    Remove-DistributionGroup -Identity $file
}