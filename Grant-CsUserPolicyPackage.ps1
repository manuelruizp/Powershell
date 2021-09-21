Import-Module MicrosoftTeams

$account = Get-Credential
Connect-MicrosoftTeams -Credential $account

$excel = New-Object -Com Excel.Application

$sheet = $excel.Workbooks.Open("D:\OneDrive - Colegio Bautista Cristiano\PowerShell Teams\Excel\policypackage.xlsx").Sheets.Item(1)

$startRow = 2
$lastRow = 27

for ($i = $startRow; $i -le $lastRow; $i++) {
    # Inicianlizando variables
    $account_name = ''
    $policy_package = ''

    # Primera columna: Cuenta del estudiante
    if ($sheet.Cells.Item($i, 1).Text -ne '') {
        $account_name = $sheet.Cells.Item($i, 1).Text
    }
    else {
        Write-Output "Falta la cuenta del estudiante en la linea $($i). Deteniendo la ejecución"
        break
    }

    # Segunda columna: ID del Team
    if ($sheet.Cells.Item($i, 2).Text -ne '') {
        $policy_package = $sheet.Cells.Item($i, 2).Text
    }
    else {
        Write-Output "Falta el ID del grupo en la linea $($i). Deteniendo la ejecución."
        break
    }
 
    if ($account_name -ne '') {
        if ($policy_package -ne '') {
            Write-Output "Asignando a la cuenta $($account_name) el paquete $($policy_package)"
            Grant-CsUserPolicyPackage -Identity $account_name -PackageName $policy_package
        }
    }

    # Pausar el programa por 1 segundo 
    # Start-Sleep -Seconds 0.5
}

Write-Output "Paquetes de seguridad asignados exitosamente"

Disconnect-MicrosoftTeams
$excel.Workbooks.Close()