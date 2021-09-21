Import-Module AzureAD

$account = Get-Credential
Connect-AzureAD -Credential $account

$excel = New-Object -Com Excel.Application

$sheet = $excel.Workbooks.Open("D:\OneDrive - Colegio Bautista Cristiano\CBC\Microsoft Teams\Powershell\Excel\users_changepassword.xlsx").Sheets.Item(1)

$startRow = 119
$lastRow = 461

$counter = 0
for ($i = $startRow; $i -le $lastRow; $i++) {
    # Primera columna: ObjectId
    if ($sheet.Cells.Item($i, 1).Text -ne '') {
        $ObjectId = $sheet.Cells.Item($i, 1).Text
    }
    else {
        Write-Output "Falta el campo ObjectId en la linea $($i). Deteniendo la ejecución."
        break
    }

    # Segunda columna: UserPrincipalName
    if ($sheet.Cells.Item($i, 2).Text -ne '') {
        $UserPrincipalName = $sheet.Cells.Item($i, 2).Text
    }
    else {
        Write-Output "Falta el campo UserPrincipalName en la linea $($i). Deteniendo la ejecución."
        break
    }

    # Tercera columna: Password
    if ($sheet.Cells.Item($i, 3).Text -ne '') {
        $pass = Write-Output $sheet.Cells.Item($i, 3).Text
        $Password = ConvertTo-SecureString $pass -AsPlainText -Force
    }
    else {
        Write-Output "Falta el campo Password en la linea $($i). Deteniendo la ejecución."
        break
    }
 
    try
    {
        Set-AzureADUserPassword -ObjectId $ObjectId -Password $Password -ForceChangePasswordNextLogin ([bool]([int]1))
        Write-Output "La clave de $($UserPrincipalName) cuyo objeto es $($ObjectId) fue modificada" 
        $counter = $counter + 1
    }
    catch
    {
        Write-Output "No se pudo cambiar la clave de $($UserPrincipalName) cuyo objeto es $($ObjectId)" 
        # Write-Host $_.Exception.Message -ForegroundColor Yellow
    }
    

    # Reiniciar variables
    $ObjectId = $UserPrincipalName = $pass = ''
    
    Start-Sleep -Seconds 1
}

Write-Output "La clave de $($counter) usuarios fueron modificadas exitosamente"
$excel.Workbooks.Close()

Disconnect-AzureAD

