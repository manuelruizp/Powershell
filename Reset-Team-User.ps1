Import-Module MicrosoftTeams

# Autenticación en el dominio
$account = Get-Credential
Connect-MicrosoftTeams -Credential $account

$excel = New-Object -Com Excel.Application
$sheet = $excel.Workbooks.Open("D:\Dropbox\CBC\Microsoft Teams\excel\reset.xlsx").Sheets.Item(1)

$startRow = 2
$lastRow = 64
$user = 'prueba@cbc.edu.do'

for ($i = $startRow; $i -le $lastRow; $i++) {
    if ($sheet.Cells.Item($i, 1).Text -ne '') {
        $groupId = $sheet.Cells.Item($i, 1).Text
        
        Write-Output "Agregando manualmente a: $($groupId)"

        Add-TeamUser -GroupId $groupId -User $user -Role Member
        Add-TeamUser -GroupId $GroupId -User $user -Role Owner
        # Start-Sleep -Seconds 0.5
        
        # Remove-TeamUser -GroupId $groupId -User $user        
    }
}

$excel.Workbooks.Close()