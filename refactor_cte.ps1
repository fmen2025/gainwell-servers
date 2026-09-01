# Extract server names and convert to VALUES clause format
$filePath = 'c:\projects\GainwellServers_20260820\output\sts_server_comparison.sql'
$content = Get-Content $filePath -Raw

# Extract all server names from the UNION ALL pattern
$pattern = "SELECT ita\.fnSnowServer\(N'([^']+)', 'Servername'\)"
$matches = [regex]::Matches($content, $pattern)

$serverNames = @()
foreach ($match in $matches) {
    $serverNames += $match.Groups[1].Value
}

Write-Host "Found $($serverNames.Count) servers"

# Build the VALUES clause
$valuesLines = @()
foreach ($name in $serverNames) {
    $valuesLines += "    (N'$name')"
}

# Join with comma+newline
$valuesClause = $valuesLines -join ",`n"

# Build the complete CTE replacement
$newCTE = @"
WITH excel_servers AS (
  SELECT ServerName FROM (VALUES
$valuesClause
  ) AS t(ServerName)
)
"@

# Find and replace the entire CTE block
$pattern = "WITH excel_servers AS \(\s+SELECT ita\.fnSnowServer.*?vmprdeuspk04', 'Servername'\) AS Servername\s+SELECT ita\.fnSnowServer\(N'vmprdeustw01'.*?\) AS Servername\s+\)"
$newContent = [regex]::Replace($content, $pattern, $newCTE, [System.Text.RegularExpressions.RegexOptions]::Singleline)

# Save the new version
$newContent | Set-Content $filePath -Encoding UTF8

Write-Host "CTE refactored successfully!"
Write-Host "Old size: $($content.Length) chars"
Write-Host "New size: $($newContent.Length) chars"
Write-Host "Reduction: $($content.Length - $newContent.Length) chars"
