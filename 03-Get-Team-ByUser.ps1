Import-Module MicrosoftTeams

# Autenticación en el dominio
$account = Get-Credential
Connect-MicrosoftTeams -Credential $account

$user = Read-Host -Prompt 'Digite el nombre del usuario para consultar sus Teams: '

Get-Team -User $user
