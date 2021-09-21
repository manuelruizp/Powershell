Import-Module MicrosoftTeams

# Autenticación en el dominio
$account = Get-Credential
Connect-MicrosoftTeams -Credential $account

$user = Read-Host "Digite el nombre de un usuario"
   
Write-Output "Consultando los equipos del usuario: $($user)"
Get-Team -User $args[0] 
