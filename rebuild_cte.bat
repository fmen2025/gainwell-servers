@echo off
REM Rebuild the SQL file with VALUES clause instead of massive UNION ALL

powershell -NoProfile -Command {
    $file = 'c:\projects\GainwellServers_20260820\output\sts_server_comparison.sql'
    $content = Get-Content $file -Raw
    
    # Extract server names from the current pattern (both old and partially new format)
    $pattern1 = "SELECT ita\.fnSnowServer\(N'([^']+)', 'Servername'\)"
    $pattern2 = "\(N'([^']+)'\)(?=\s*,)"
    
    $names1 = [regex]::Matches($content, $pattern1) | ForEach-Object { $_.Groups[1].Value }
    $names2 = [regex]::Matches($content, $pattern2) | ForEach-Object { $_.Groups[1].Value }
    
    $allNames = @($names1 + $names2) | Sort-Object -Unique
    
    Write-Host "Found $($allNames.Count) unique server names"
    
    # Build VALUES clause
    $vals = @()
    foreach ($name in $allNames) {
        $vals += "    (N'$name')"
    }
    $valuesStr = $vals -join ",`n"
    
    # New CTE structure
    $newCTE = @"
WITH excel_servers AS (
  SELECT ServerName FROM (VALUES
$valuesStr
  ) AS t(ServerName)
)
"@
    
    # Replace the entire CTE (greedy match from WITH to last closing paren)
    $newContent = $content -replace "WITH excel_servers AS \(\s+SELECT.*?\s+\)", $newCTE, [System.Text.RegularExpressions.RegexOptions]::Singleline
    
    if ($newContent -ne $content) {
        $newContent | Set-Content $file -Encoding UTF8
        Write-Host "File updated successfully!"
        Write-Host "Old: $($content.Length) | New: $($newContent.Length) | Saved: $(($content.Length - $newContent.Length) / 1024)KB"
    } else {
        Write-Host "No changes made - pattern didn't match"
    }
}
