# Import-Module MicrosoftTeams

# $account = Get-Credential
# Connect-MicrosoftTeams -Credential $account

$excel = New-Object -Com Excel.Application

$sheet = $excel.Workbooks.Open("D:\Dropbox\CBC\Microsoft Teams\excel\Plantillas.xlsx").Sheets.Item(4)

$startRow = 2
$lastRow = 5

for ($i = $startRow; $i -le $lastRow; $i++) {

    # Primera columna: TEAM_NAME
    if ($sheet.Cells.Item($i, 1).Text -ne '') {
        $name = $sheet.Cells.Item($i, 1).Text
    }
    else {
        Write-Output "Falta el nombre requerida para crear un Team en la linea $($i). Deteniendo la ejecución"
        break
    }

    # Segunda columna: ALIAS
    if ($sheet.Cells.Item($i, 2).Text -ne '') {
        $alias = $sheet.Cells.Item($i, 2).Text
    }
    else {
        Write-Output "Falta el alias requerida para crear un Team en la linea $($i). Deteniendo la ejecución."
        break
    }

    # Tercera columna: COORDINADOR
    if ($sheet.Cells.Item($i, 3).Text -ne '') {
        $coordinator = Write-Output $sheet.Cells.Item($i, 3).Text
    }
    else {
        Write-Output "Falta la cuenta del coordinador requerida para crear un Team en la linea $($i). Deteniendo la ejecución."
        break
    }
 
    # Cuarta columna: MAESTRO (NO REQUERIDO)
    if ($sheet.Cells.Item($i, 4).Text -ne '') {
        $teacher = Write-Output $sheet.Cells.Item($i, 4).Text
    }

    # Quinta columna: ASISTENTE (NO REQUERIDO)
    if ($sheet.Cells.Item($i, 5).Text -ne '') {
        $assistant = Write-Output $sheet.Cells.Item($i, 5).Text
    }

    $new_team_id = ( New-Team -DisplayName $name -Owner $coordinator -Alias $alias -Template "EDU_Class" )

    if ($teacher -ne '') {
        Add-TeamUser -GroupId $new_team_id.GroupId -User $teacher -Role Member
        Add-TeamUser -GroupId $new_team_id.GroupId -User $teacher -Role Owner
    }

    if ($assistant -ne '') {
        Add-TeamUser -GroupId $new_team_id.GroupId -User $assistant -Role Member
        Add-TeamUser -GroupId $new_team_id.GroupId -User $assistant -Role Owner
    }
}

$excel.Workbooks.Close()