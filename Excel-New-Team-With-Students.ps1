Import-Module MicrosoftTeams

$account = Get-Credential
Connect-MicrosoftTeams -Credential $account

$excel = New-Object -Com Excel.Application

$sheet = $excel.Workbooks.Open("D:\OneDrive - Colegio Bautista Cristiano\CBC\Microsoft Teams\Powershell\Excel\DAD.xlsx").Sheets.Item(1)

$startRow = 2
$lastRow = 28

for ($i = $startRow; $i -le $lastRow; $i++) {
    # Primera columna: TEAM_NAME
    if ($sheet.Cells.Item($i, 1).Text -ne '') {
        $TeamName = $sheet.Cells.Item($i, 1).Text
    }
    else {
        Write-Output "Falta el nombre requerida para crear un Team en la linea $($i). Deteniendo la ejecución"
        break
    }

    # Segunda columna: TEAM_ALIAS
    if ($sheet.Cells.Item($i, 2).Text -ne '') {
        $TeamAlias = $sheet.Cells.Item($i, 2).Text
    }
    else {
        Write-Output "Falta el alias para crear un Team en la linea $($i). Deteniendo la ejecución"
        break
    }
    
    # Tercera columna: Creador del Team
    if ($sheet.Cells.Item($i, 3).Text -ne '') {
        $Coordinator = $sheet.Cells.Item($i, 3).Text
    }
    else {
        Write-Output "Falta el creador del Team en la linea $($i). Deteniendo la ejecución."
        break
    }

    # Cuarta columna: Arreglo de Maestros
    if ($sheet.Cells.Item($i, 4).Text -ne '') {
        $OwnersArr = ($sheet.Cells.Item($i, 4).Text).Split(',')
    }
    else {
        Write-Output "Falta los propietarios requeridos para crear un Team en la linea $($i). Deteniendo la ejecución."
        break
    }

    # Quinta columna: Arreglo de Estudiantes
    if ($sheet.Cells.Item($i, 5).Text -ne '') {
        $MembersArr = ($sheet.Cells.Item($i, 5).Text).Split(',')
    }
    else {
        Write-Output "Falta los estudiantes requeridos para crear un Team en la linea $($i). Deteniendo la ejecución."
        break
    }

    Write-Output "Creando el Team: $($TeamName) coordinado por $($Coordinator)"
    $new_team = ( New-Team -DisplayName $TeamName -Owner $Coordinator -Alias $TeamAlias -Template "EDU_Class")
    
    foreach ($Owner in $OwnersArr) {
        Write-Output "Insertando el maestro $($Owner) al Team $($new_team.GroupId)"
        Add-TeamUser -GroupId $new_team.GroupId -User $Owner -Role Member
        Add-TeamUser -GroupId $new_team.GroupId -User $Owner -Role Owner
    }

    foreach ($Member in $MembersArr) {
        Write-Output "Insertando el estudiante $($Member) al Team $($new_team.GroupId)"
        Add-TeamUser -GroupId $new_team.GroupId -User $Member -Role Member
    }
}

Write-Output "Teams creados."

$excel.Workbooks.Close()
Disconnect-MicrosoftTeams