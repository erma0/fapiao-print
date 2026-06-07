$files = @(
    'd:\test\fapiao\src\app.js',
    'd:\test\fapiao\src\print.js',
    'd:\test\fapiao\src\ocr.js'
)

foreach ($file in $files) {
    $content = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)

    # Step 1: Replace invoke( with __api.call( — all function calls
    $content = $content -replace 'invoke\(', '__api.call('

    # Step 2: Replace 'isTauri && invoke' with 'isTauri && __api' (conditional checks)
    $content = $content -replace 'isTauri && invoke', 'isTauri && __api'

    # Step 3: Replace '!invoke' with '!__api' (negation checks)
    $content = $content -replace '!invoke', '!__api'

    # Step 4: Remove the 'var invoke = ...' declaration line (only in app.js)
    $content = $content -replace 'var invoke\s+= isTauri \? window\.__TAURI_INTERNALS__\.invoke : null;\r?\n', ''

    # Step 5: Update dependency comments
    $content = $content -replace '// Dependencies \(global\): isTauri, invoke,', '// Dependencies (global): isTauri, __api,'
    $content = $content -replace '// Dependencies \(global\): invoke, isTauri,', '// Dependencies (global): __api, isTauri,'

    [System.IO.File]::WriteAllText($file, $content, (New-Object System.Text.UTF8Encoding $false))
    Write-Host "Replaced: $file"
}

Write-Host "All files replaced successfully"
