Import-Module MicrosoftTeams
$account = Get-Credential
Connect-MicrosoftTeams -Credential $account