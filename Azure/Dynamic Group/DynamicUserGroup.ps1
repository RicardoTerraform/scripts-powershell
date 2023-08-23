#se o -membershipRule não funcionar, correr estes dois comandos
#Uninstall-Module -Name AzureAD
#Install-Module -Name AzureADPreview -RequiredVersion 2.0.2.129

Connect-AzureAD

$Groups = Import-Csv -Path "C:\Users\DynamicGroups.csv"
$dynamicGroupTypeString = "DynamicMembership"

foreach($Group in $Groups)
{
    New-AzureADMSGroup -DisplayName $Group.DisplayName -Description $Group.Description -MailEnabled $False -MailNickName "group" -SecurityEnabled $True -membershipRule $Group.MembershipRule -GroupTypes $dynamicGroupTypeString -MembershipRuleProcessingState "On"
}

 