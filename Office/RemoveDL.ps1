Connect-ExchangeOnline

$file_path= Get-Content C:\Users\dados.txt

foreach($file in $file_path)
{
    Remove-DistributionGroup -Identity $file
}