
$excel = New-Object -Com Excel.Application
$wb = $excel.Workbooks.Open("D:\Dropbox\CBC\Microsoft Teams\excel\Inicial.xlsx")

$sheet = $wb.Sheets.Item(1)

for ($i = 2; $i -le 77; $i++) {
    if ($sheet.Cells.Item($i, 1).Text -ne '') {
        $TEAM_NAME = $sheet.Cells.Item($i, 1).Text
    }

    if ($sheet.Cells.Item($i, 5).Text -ne '') {
        $ALIAS = Write-Output $sheet.Cells.Item($i, 5).Text
    }

    if ($sheet.Cells.Item($i, 2).Text -ne '') {
        $TEACHER_ACCOUNT = Write-Output $sheet.Cells.Item($i, 2).Text
    }

    if ($sheet.Cells.Item($i, 3).Text -ne '') {
        $ASSISTANT_ACCOUNT = Write-Output $sheet.Cells.Item($i, 3).Text
    }

    if ($sheet.Cells.Item($i, 4).Text -ne '') {
        $COORDINATOR_ACCOUNT = Write-Output $sheet.Cells.Item($i, 4).Text
    }

    Write-Output "$($TEAM_NAME) $($TEACHER_ACCOUNT) $($ASSISTANT_ACCOUNT) $($COORDINATOR_ACCOUNT)"

}

$excel.Workbooks.Close()