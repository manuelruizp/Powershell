# Script para eliminar un listado de estudiantes con sus licencias
# Import-Module MicrosoftTeams

$account = Get-Credential
Connect-MsolService -Credential $account
Connect-AzureAD -Credential $account

$Students = @('albert.cruz@est.cbc.edu.do', 'edward.gil@est.cbc.edu.do', 'hamlet.fernandez@est.cbc.edu.do', 'hannah.cruz@est.cbc.edu.do', 'sarah.fernandez@est.cbc.edu.do', 'weldris.gil@est.cbc.edu.do')

for ($i = 0; $i -lt $Students.count; $i++)
{
   Write-Output "Eliminando licencia Office365 Student A1 de $($Students[$i])"
   Set-MsolUserLicense -UserPrincipalName $Students[$i] -RemoveLicenses "cbcedudo:STANDARDWOFFPACK_STUDENT"
}

for ($i = 0; $i -lt $Students.count; $i++)
{
   Write-Output "Eliminando estudiante $($Students[$i])"
   Remove-AzureADUser -ObjectID $($Students[$i])
}

Write-Output "Estudiantes eliminados" 

Disconnect-AzureAD
