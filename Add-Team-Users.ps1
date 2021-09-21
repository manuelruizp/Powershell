# Script para eliminar un listado de estudiantes con sus licencias
# Import-Module MicrosoftTeams

# Autenticación en el dominio
$account = Get-Credential
Connect-MicrosoftTeams -Credential $account

# Usuarios enumerado por coma
$Users = @('email@cbc.edu.do')

for ($i = 0; $i -lt $Users.count; $i++)
{
   Write-Output "Agregando usuario al equipo: $($Users[$i])"
   Add-TeamUser -GroupId "ec5e2b89-c415-415b-bf93-294cb69b9d2a" -User $($Users[$i]) -Role Member
}
