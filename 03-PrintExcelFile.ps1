
$excel = New-Object -Com Excel.Application
$wb = $excel.Workbooks.Open("D:\PowerShell\Inicial.xls")

# $startRow = 2
# $startColumn = 1
# $endRow = 77
# $endColumn = 3

$sheet = $wb.Sheets.Item(1)
for ($i = 1; $i -le 77; $i++) {
    Write-Output  $sheet.Cells.Item($i, 1).Text
}

$excel.Workbooks.Close()