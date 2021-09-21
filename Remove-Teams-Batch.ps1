Import-Module MicrosoftTeams

# Autenticación en el dominio
$account = Get-Credential
Connect-MicrosoftTeams -Credential $account

# El id de cada Teams a eliminar dividido por coma
$deleteTeams = @('f4020d13-dd57-4ede-9799-100842584a08', '09737a76-06b3-4afb-a3ba-bf727fd39c0b', 'f6c2fe49-5d39-4f91-837a-3750442b3c7a')

for ($i = 0; $i -lt $deleteTeams.count; $i++)
{
   Write-Output "Eliminando el teams con ID: $($deleteTeams[$i])"
   Remove-Team -GroupId $($deleteTeams[$i])
}
