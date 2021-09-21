Import-Module AzureAD

$account = Get-Credential
Connect-AzureAD -Credential $account

$excel = New-Object -Com Excel.Application

$sheet = $excel.Workbooks.Open("D:\OneDrive - Colegio Bautista Cristiano\CBC\Microsoft Teams\Powershell\Excel\users.xlsx").Sheets.Item(1)

$startRow = 2
$lastRow = 63

$PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile

$counter = 0

for ($i = $startRow; $i -le $lastRow; $i++) {
    # Primera columna: DisplayName
    if ($sheet.Cells.Item($i, 1).Text -ne '') {
        $DisplayName = $sheet.Cells.Item($i, 1).Text
    }
    else {
        Write-Output "Falta el campo DisplayName en la linea $($i). Deteniendo la ejecución."
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

    # Tercera columna: FirstName
    if ($sheet.Cells.Item($i, 3).Text -ne '') {
        $FirstName = Write-Output $sheet.Cells.Item($i, 3).Text
    }
    else {
        Write-Output "Falta el campo FirstName en la linea $($i). Deteniendo la ejecución."
        break
    }
 
    # Cuarta columna: LastName
    if ($sheet.Cells.Item($i, 4).Text -ne '') {
        $LastName = Write-Output $sheet.Cells.Item($i, 4).Text
    }
    else {
        Write-Output "Falta el campo LastName en la linea $($i). Deteniendo la ejecución."
        break
    }

    # Quinta columna: LicenseAssignment
    if ($sheet.Cells.Item($i, 5).Text -ne '') {
        $LicenseAssignment = Write-Output $sheet.Cells.Item($i, 5).Text
    }
    else {
        Write-Output "Falta el campo LicenseAssignment en la linea $($i). Deteniendo la ejecución."
        break
    }

    # Sexta columna: Password
    if ($sheet.Cells.Item($i, 6).Text -ne '') {
        $Password = Write-Output $sheet.Cells.Item($i, 6).Text
        $PasswordProfile.Password = $Password
    }
    else {
        Write-Output "Falta el campo Password en la linea $($i). Deteniendo la ejecución."
        break
    }

    # Sexta columna: MailNickname
    if ($sheet.Cells.Item($i, 7).Text -ne '') {
        $MailNickname = Write-Output $sheet.Cells.Item($i, 7).Text
    }
    else {
        Write-Output "Falta el campo MailNickname en la linea $($i). Deteniendo la ejecución."
        break
    }
 
    Write-Output "Creando nuevo usuario con el nombre $($DisplayName) con licencia $($LicenseAssignment)" 

    New-AzureADUser -DisplayName $DisplayName -GivenName $FirstName -SurName $LastName -UserPrincipalName $UserPrincipalName -PasswordProfile $PasswordProfile -AccountEnabled $true -MailNickName $MailNickname

    # Reiniciar variables
    $DisplayName = $FirstName = $LastName = $UserPrincipalName = $Password = $MailNickname = ''

    $counter = $counter + 1
}

Write-Output "$($counter) usuarios creados exitosamente"
$excel.Workbooks.Close()

Disconnect-AzureAD

