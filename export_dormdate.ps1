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
        if ($count -gt 20) { break }
        $cells = $row.Cells
        $values = @()
        foreach ($cell in $cells) {
            $text = $cell.Text
            if (-not $text) { $text = $cell.Value2 }
            $values += [string]$text
        }
        $data += ,$values
        $count++
    }
    $workbook.Close($false)
    $excel.Quit()
    $data | ConvertTo-Json -Depth 5 -Compress | Set-Content "d:\AI\dorm-manager\export_dormdate.json" -Encoding UTF8
    Write-Host "Success"
} catch {
    Write-Host "Error: $($_.Exception.Message)"
    if ($excel) { $excel.Quit() }
}
