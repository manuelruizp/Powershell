# Script para generar un Equipo en Microsoft Teams para Clases del centro educativo
Import-Module MicrosoftTeams

$account = Get-Credential
Connect-MicrosoftTeams -Credential $account

$DisplayName = "CAPILLA PRIMARIA"

$Alias = "CapillaPrimaria"
$Template = "EDU_Class"

# Dueños del Team
$Owner = "teamsmgr@cbc.edu.do"

# Arreglos para asistentes del profesor y los estudiantes
$CoOwners = @('s.peralta@cbc.edu.do', 'j.brito@cbc.edu.do')

$Students = @('ana.amiama@est.cbc.edu.do', 'shadday.arbaje@est.cbc.edu.do', 'sarah.chavez@est.cbc.edu.do', 'gia.frias@est.cbc.edu.do', 'lis.lara@est.cbc.edu.do', 'noa.lara@est.cbc.edu.do', 'amelia.lora@est.cbc.edu.do', 'daniella.madera@est.cbc.edu.do', 'amanda.mueses@est.cbc.edu.do', 'sofia.ramirez@est.cbc.edu.do', 'sofia.rodriguez@est.cbc.edu.do', 'rebeca.santiago@est.cbc.edu.do', 'vivian.sosa@est.cbc.edu.do', 'isaac.bueno@est.cbc.edu.do', 'oliver.espinosa@est.cbc.edu.do', 'alejandro.estrella@est.cbc.edu.do', 'jacob.frias@est.cbc.edu.do', 'sebastian.guzman@est.cbc.edu.do', 'carlos.mesa@est.cbc.edu.do', 'emmanuel.nanita@est.cbc.edu.do', 'lucas.rodriguez@est.cbc.edu.do', 'diego.rojas@est.cbc.edu.do', 'jedidah.auguiste@est.cbc.edu.do', 'hannah.cruz@est.cbc.edu.do', 'camila.diaz@est.cbc.edu.do', 'sara.martinez@est.cbc.edu.do', 'sarah.minaya@est.cbc.edu.do', 'lia.puig@est.cbc.edu.do', 'miah.sanchez@est.cbc.edu.do', 'dariana.socias@est.cbc.edu.do', 'ana.tejeda@est.cbc.edu.do', 'emmanuel.abreu@est.cbc.edu.do', 'gabriel.delossantos@est.cbc.edu.do', 'joan.feliz@est.cbc.edu.do', 'hamlet.fernandez@est.cbc.edu.do', 'gariel.lopez@est.cbc.edu.do', 'christopher.luciano@est.cbc.edu.do', 'manuel.nunez@est.cbc.edu.do', 'amaury.perdomo@est.cbc.edu.do', 'miguel.rodriguez@est.cbc.edu.do', 'carlos.santana@est.cbc.edu.do', 'jean.tempestti@est.cbc.edu.do', 'emely.amparo@est.cbc.edu.do', 'zoe.disla@est.cbc.edu.do', 'amalia.mendoza@est.cbc.edu.do', 'laura.ortiz@est.cbc.edu.do', 'ana.rivas@est.cbc.edu.do', 'valeria.sanchez@est.cbc.edu.do', 'zoe.tineo@est.cbc.edu.do', 'roger.bournigal@est.cbc.edu.do', 'brandon.burgos@est.cbc.edu.do', 'nathan.cabral@est.cbc.edu.do', 'jean.cartagena@est.cbc.edu.do', 'daniel.collado@est.cbc.edu.do', 'emilio.lopez@est.cbc.edu.do', 'samuel.martinez@est.cbc.edu.do', 'pedro.molina@est.cbc.edu.do', 'aiser.ramirez@est.cbc.edu.do', 'yody.reyes@est.cbc.edu.do', 'joel.roque@est.cbc.edu.do', 'margaret.cadena@est.cbc.edu.do', 'mayrham.colon@est.cbc.edu.do', 'marcela.dicen@est.cbc.edu.do', 'esther.garcia@est.cbc.edu.do', 'sarah.gonzalez@est.cbc.edu.do', 'zoe.morel@est.cbc.edu.do', 'lisa.nanita@est.cbc.edu.do', 'gabriela.nunez@est.cbc.edu.do', 'mia.polanco@est.cbc.edu.do', 'karlamichelle.reyes@est.cbc.edu.do', 'karlabeatriz.reyes@est.cbc.edu.do', 'yadelin.rivera@est.cbc.edu.do', 'noel.arias@est.cbc.edu.do', 'sebastian.baez@est.cbc.edu.do', 'nelson.caballero@est.cbc.edu.do', 'isaac.encarnacion@est.cbc.edu.do', 'maicol.gonzalez@est.cbc.edu.do', 'armando.lora@est.cbc.edu.do', 'jose.montero@est.cbc.edu.do', 'nathan.perez@est.cbc.edu.do', 'sebastian.rodriguez@est.cbc.edu.do', 'eric.sanchez@est.cbc.edu.do', 'natasha.bernardino@est.cbc.edu.do', 'melanie.cleto@est.cbc.edu.do', 'sarai.coste@est.cbc.edu.do', 'rachell.cruz@est.cbc.edu.do', 'noely.cuello@est.cbc.edu.do', 'abigail.duran@est.cbc.edu.do', 'habana.matos@est.cbc.edu.do', 'perla.mesa@est.cbc.edu.do', 'micaela.rodriguez@est.cbc.edu.do', 'crystal.santana@est.cbc.edu.do', 'maryam.vicente@est.cbc.edu.do', 'sebastian.baezferreira@est.cbc.edu.do', 'miguel.guzman@est.cbc.edu.do', 'lucas.hernandez@est.cbc.edu.do', 'jireh.matos@est.cbc.edu.do', 'joaquin.mercedes@est.cbc.edu.do', 'benjamin.nadal@est.cbc.edu.do', 'juan.russo@est.cbc.edu.do', 'angel.soto@est.cbc.edu.do', 'carlos.tavarez@est.cbc.edu.do')

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