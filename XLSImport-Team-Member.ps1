Import-Module MicrosoftTeams

$account = Get-Credential
Connect-MicrosoftTeams -Credential $account

$excel = New-Object -Com Excel.Application

$sheet = $excel.Workbooks.Open("C:\Users\Manuel Ruiz\OneDrive - Colegio Bautista Cristiano\PowerShell Teams\Excel\transactions.xlsx").Sheets.Item(1)

$startRow = 2
$lastRow = 128

for ($i = $startRow; $i -le $lastRow; $i++) {
    # Inicianlizando variables
    $group_id = ''
    $account_name = ''

    # Segunda columna: ID del Team
    if ($sheet.Cells.Item($i, 2).Text -ne '') {
        $group_id = $sheet.Cells.Item($i, 2).Text
    }
    else {
        Write-Output "Falta el ID del grupo en la linea $($i). Deteniendo la ejecución."
        break
    }

    
    # Quinta columna: Cuenta del estudiante
    if ($sheet.Cells.Item($i, 5).Text -ne '') {
        $account_name = $sheet.Cells.Item($i, 5).Text
    }
    else {
        Write-Output "Falta la cuenta del estudiante en la linea $($i). Deteniendo la ejecución"
        break
    }

    if ($account_name -ne '') {
        if ($group_id -ne '') {
            Write-Output "Insertando el estudiante con cuenta $($account_name) al team $($group_id)"            
            Add-TeamUser -GroupId $group_id -User $account_name -Role Member
        }
    }

    # Pausar el programa por 1 segundo 
    # Start-Sleep -Seconds 0.5
}

Write-Output "Estudiantes insertados exitosamente a los Teams"

Disconnect-MicrosoftTeams
$excel.Workbooks.Close()