$f = Get-Content 'c:\Users\flavi\Downloads\Retail_Store_Sales_Data\dashboard_retail_store_part1.html' -Encoding UTF8
$em = [char]0x2014
for ($idx=0; $idx -lt $f.Count; $idx++) {
    if ($f[$idx].Contains($em)) {
        $line = $f[$idx]
        if ($line.Length -gt 150) { $line = $line.Substring(0, 150) }
        Write-Host ("{0}: {1}" -f ($idx+1), $line)
    }
}
