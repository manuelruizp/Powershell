# Script para generar un Equipo en Microsoft Teams para Clases del centro educativo
Import-Module MicrosoftTeams

$account = Get-Credential
Connect-MicrosoftTeams -Credential $account

$DisplayName = "CAPILLA PRIMARIA"

$Alias = "CapillaPrimaria"
$Template = "EDU_Class"

# Dueños del Team
$Owner = "dueno@cbc.edu.do"

# Arreglos para asistentes del profesor y los estudiantes
$CoOwners = @()

# Arreglo de estudiantes dividido por comma
$Students = @()

$new_team_id = ( New-Team -DisplayName $DisplayName -Description $Description -Owner $Owner -Alias $Alias -Template $Template )

for ($i = 0; $i -lt $CoOwners.count; $i++)
{
    Write-Output "Agregando la cuenta de $($CoOwners[$i]) como maestro"
    Add-TeamUser -GroupId $new_team_id.GroupId -User $CoOwners[$i] -Role Member
    Add-TeamUser -GroupId $new_team_id.GroupId -User $CoOwners[$i] -Role Owner
}

for ($i = 0; $i -lt $Students.count; $i++)
{
    Write-Output "Agregando la cuenta de $($Students[$i]) como estudiante"
    Add-TeamUser -GroupId $new_team_id.GroupId -User $Students[$i] -Role Member
}

Write-Output "Nuevo Team creado: $($new_team_id)" 

Disconnect-MicrosoftTeams