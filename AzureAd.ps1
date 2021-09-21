
Import-Module AzureAD

$account = Get-Credential
Connect-AzureAD -Credential $account

Get-AzureADUser -ObjectId "aaron.ramirez@est.cbc.edu.do"
