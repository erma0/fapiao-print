$p1 = "C:\Users\Administrator\Downloads\【K9用车-177.86元-1个行程】高德打车电子发票.pdf"
Write-Host "=== 高德打车发票 ==="
Write-Host "Exists: $([System.IO.File]::Exists($p1))"
$b1 = [System.IO.File]::ReadAllBytes($p1)
Write-Host "Size: $($b1.Length) bytes"
Write-Host "Header: $($b1[0..19] -join ' ')"

$p2 = "C:\Users\Administrator\Downloads\25329116804007140998.pdf"
Write-Host ""
Write-Host "=== 火车票PDF ==="
Write-Host "Exists: $([System.IO.File]::Exists($p2))"
$b2 = [System.IO.File]::ReadAllBytes($p2)
Write-Host "Size: $($b2.Length) bytes"
Write-Host "Header: $($b2[0..19] -join ' ')"
