# Script para generar un Equipo en Microsoft Teams para Clases del centro educativo
Import-Module MicrosoftTeams

$account = Get-Credential
Connect-MicrosoftTeams -Credential $account

$DisplayName = "PowerShell Team Type Class"

# Descripción es opcional en la creación del Team (ahora mismo será igual al título)
$Description = $DisplayName
$Alias = "PowerShellTeamTypeClass"
$Template = "EDU_Class"

# Dueños del Team
$Owner = "prueba@domain.com"

# Arreglos para asistentes del profesor y los estudiantes
$CoOwners = @('prueba1@domain.com', 'prueba2@domain.com')
$Students = @('estudiante1@dominio.com', 'estudiante2@domain.com')

$new_team_id = ( New-Team -DisplayName $DisplayName -Description $Description -Owner $Owner -Alias $Alias -Template $Template )

for ($i = 0; $i -lt $CoOwners.count; $i++)
{
    Add-TeamUser -GroupId $new_team_id.GroupId -User $CoOwners[$i] -Role Member
    Add-TeamUser -GroupId $new_team_id.GroupId -User $CoOwners[$i] -Role Owner
}

for ($i = 0; $i -lt $Students.count; $i++)
{
    Add-TeamUser -GroupId $new_team_id.GroupId -User $Students[$i] -Role Member
    Add-TeamUser -GroupId $new_team_id.GroupId -User $Students[$i] -Role Owner
}

Write-Output "Nuevo Team creado: $($new_team_id)" 

Disconnect-MicrosoftTeams