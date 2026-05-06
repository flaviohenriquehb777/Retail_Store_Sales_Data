# Extract embedded JSON data from original file
$orig = Get-Content "c:\Users\flavi\Downloads\Retail_Store_Sales_Data\dashboard_retail_store.html" -Raw
$match = [regex]::Match($orig, '(<script id="embeddedData"[^>]*>)(.*?)(</script>)', [System.Text.RegularExpressions.RegexOptions]::Singleline)
$jsonData = $match.Groups[2].Value

# Read template parts
$part1 = Get-Content "c:\Users\flavi\Downloads\Retail_Store_Sales_Data\_part1.html" -Raw
$part2 = Get-Content "c:\Users\flavi\Downloads\Retail_Store_Sales_Data\_part2.html" -Raw

# Combine
$final = $part1 + $jsonData + $part2
[System.IO.File]::WriteAllText("c:\Users\flavi\Downloads\Retail_Store_Sales_Data\dashboard_retail_store.html", $final, [System.Text.Encoding]::UTF8)

# Cleanup
Remove-Item "c:\Users\flavi\Downloads\Retail_Store_Sales_Data\_part1.html" -Force
Remove-Item "c:\Users\flavi\Downloads\Retail_Store_Sales_Data\_part2.html" -Force
Remove-Item "c:\Users\flavi\Downloads\Retail_Store_Sales_Data\rebuild.ps1" -Force

Write-Host "Dashboard rebuilt successfully!"
