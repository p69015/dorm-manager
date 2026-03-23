$ErrorActionPreference = "Stop"
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open("d:\AI\dorm-manager\dormdate.xlsx")
    $sheet = $workbook.Sheets.Item(1)
    $rows = $sheet.UsedRange.Rows
    $data = @()
    $count = 0
    foreach ($row in $rows) {
        if ($count -gt 5) { break }
        $cells = $row.Cells
        $values = @()
        foreach ($cell in $cells) {
            $text = $cell.Text
            if (-not $text) { $text = $cell.Value2 }
            $values += [string]$text
        }
        $data += ($values -join "|")
        $count++
    }
    $workbook.Close($false)
    $excel.Quit()
    Write-Host ($data -join "`n")
} catch {
    Write-Host "Error: $($_.Exception.Message)"
    if ($excel) { $excel.Quit() }
}
