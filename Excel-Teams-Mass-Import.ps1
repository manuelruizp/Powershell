Import-Module MicrosoftTeams

$account = Get-Credential
Connect-MicrosoftTeams -Credential $account

$excel = New-Object -Com Excel.Application

$sheet = $excel.Workbooks.Open("D:\OneDrive - Colegio Bautista Cristiano\CBC\Microsoft Teams\Powershell\Excel\teams.xlsx").Sheets.Item(1)

$startRow = 2
$lastRow = 216

for ($i = $startRow; $i -le $lastRow; $i++) {
    # Primera columna: Displayname
    if ($sheet.Cells.Item($i, 1).Text -ne '') {
        $name = $sheet.Cells.Item($i, 1).Text
    }
    else {
        Write-Output "Falta el nombre requerida para crear un Team en la linea $($i). Deteniendo la ejecución"
        break
    }

    # Segunda columna: MailNickname
    if ($sheet.Cells.Item($i, 2).Text -ne '') {
        $MailNickname = $sheet.Cells.Item($i, 2).Text
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
        $teachers = Write-Output $sheet.Cells.Item($i, 4).Text
    }else {
        Write-Output "Falta el listado de maestros en la linea $($i). Deteniendo la ejecución."
        break
    }
 
    Write-Output "Nuevo Teams creado $($name) coordinado por $($coordinator)..."

    $new_team_id = ( New-Team -MailNickname $MailNickname -displayname $name -Owner $coordinator -Template "EDU_Class" )

    $teachersArr = $teachers.Split(",")
    foreach ($teacher in $teachersArr)
    {
        Write-Output "Agregando $($teacher) al Teams $($name) con rol de propietario..."
        Add-TeamUser -GroupId $new_team_id.GroupId -User $teacher -Role Member
        Add-TeamUser -GroupId $new_team_id.GroupId -User $teacher -Role Owner
    }

    # Reiniciar variables
    $name = ''
    $MailNickname = ''
    $coordinator = ''
    $teachers = ''

    # Pausar el programa por 1 segundo 
    Start-Sleep -Seconds 1
}

Write-Output "Creacion de teams completado"

Disconnect-MicrosoftTeams

$excel.Workbooks.Close()