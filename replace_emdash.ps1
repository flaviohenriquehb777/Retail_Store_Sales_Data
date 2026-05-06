$path = 'c:\Users\flavi\Downloads\Retail_Store_Sales_Data\dashboard_retail_store_part1.html'
$content = Get-Content $path -Raw -Encoding UTF8
$em = [char]0x2014
$newContent = $content.Replace($em, '-')
[System.IO.File]::WriteAllText($path, $newContent, [System.Text.Encoding]::UTF8)
Write-Host 'Replacement complete.'
