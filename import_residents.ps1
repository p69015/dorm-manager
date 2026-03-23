$ErrorActionPreference = "Stop"
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open("d:\AI\dorm-manager\dormdate.xlsx")
    $sheet = $workbook.Sheets.Item(1)
    $rows = $sheet.UsedRange.Rows
    $data = @()
    $isFirst = $true

    foreach ($row in $rows) {
        $cells = $row.Cells
        $values = @()
        foreach ($cell in $cells) {
            $text = $cell.Text
            if (-not $text) { $text = $cell.Value2 }
            $values += [string]$text
        }
        if ($isFirst) {
            $isFirst = $false
        } else {
            if (-not [string]::IsNullOrWhiteSpace($values[0])) {
                $data += ,$values
            }
        }
    }
    $workbook.Close($false)
    $excel.Quit()

    $json = $data | ConvertTo-Json -Depth 5 -Compress
    Set-Content "d:\AI\dorm-manager\imported_residents.js" -Value ("const IMPORTED_RESIDENTS = " + $json + ";") -Encoding UTF8
    Write-Host "Success"
} catch {
    Write-Host "Error: $($_.Exception.Message)"
    if ($excel) { $excel.Quit() }
}
