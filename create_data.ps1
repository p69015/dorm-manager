$json = Get-Content "d:\AI\dorm-manager\bedno.json" -Raw -Encoding UTF8
$jsContent = "const RAW_BED_DATA = " + $json + ";"
Set-Content "d:\AI\dorm-manager\data.js" -Value $jsContent -Encoding UTF8
