# Analyze System column from Excel against all sts.server columns

$sqlQuery = @"
-- Find matches between Excel System column and sts.server columns
DECLARE @SystemValues TABLE (SystemValue NVARCHAR(255))

-- Insert distinct System values from ita.snowserver
INSERT INTO @SystemValues
SELECT DISTINCT System FROM ita.snowserver WHERE System IS NOT NULL

DECLARE @Results TABLE (
    ExcelSystemValue NVARCHAR(255),
    sts_ServerColumn NVARCHAR(100),
    MatchCount INT,
    MatchPercentage DECIMAL(5,2)
)

DECLARE @TotalExcelSystems INT = (SELECT COUNT(DISTINCT System) FROM ita.snowserver WHERE System IS NOT NULL)

-- Check System column
INSERT INTO @Results
SELECT TOP 20
    s.System,
    'sts.server.System' AS ColumnName,
    COUNT(*) AS Matches,
    CAST(COUNT(*) * 100.0 / @TotalExcelSystems AS DECIMAL(5,2)) AS MatchPct
FROM ita.snowserver s
INNER JOIN sts.server sr ON LOWER(sr.System) = LOWER(s.System)
WHERE s.System IS NOT NULL
GROUP BY s.System
ORDER BY COUNT(*) DESC

-- Check ServerDescription column
INSERT INTO @Results
SELECT TOP 20
    s.System,
    'sts.server.ServerDescription' AS ColumnName,
    COUNT(*) AS Matches,
    CAST(COUNT(*) * 100.0 / @TotalExcelSystems AS DECIMAL(5,2)) AS MatchPct
FROM ita.snowserver s
INNER JOIN sts.server sr ON LOWER(sr.ServerDescription) = LOWER(s.System)
WHERE s.System IS NOT NULL
GROUP BY s.System
ORDER BY COUNT(*) DESC

-- Check ServerType column
INSERT INTO @Results
SELECT TOP 20
    s.System,
    'sts.server.ServerType' AS ColumnName,
    COUNT(*) AS Matches,
    CAST(COUNT(*) * 100.0 / @TotalExcelSystems AS DECIMAL(5,2)) AS MatchPct
FROM ita.snowserver s
INNER JOIN sts.server sr ON LOWER(sr.ServerType) = LOWER(s.System)
WHERE s.System IS NOT NULL
GROUP BY s.System
ORDER BY COUNT(*) DESC

-- Check PrimaryFunction column
INSERT INTO @Results
SELECT TOP 20
    s.System,
    'sts.server.PrimaryFunction' AS ColumnName,
    COUNT(*) AS Matches,
    CAST(COUNT(*) * 100.0 / @TotalExcelSystems AS DECIMAL(5,2)) AS MatchPct
FROM ita.snowserver s
INNER JOIN sts.server sr ON LOWER(sr.PrimaryFunction) = LOWER(s.System)
WHERE s.System IS NOT NULL
GROUP BY s.System
ORDER BY COUNT(*) DESC

-- Check ApplicationName column
INSERT INTO @Results
SELECT TOP 20
    s.System,
    'sts.server.ApplicationName' AS ColumnName,
    COUNT(*) AS Matches,
    CAST(COUNT(*) * 100.0 / @TotalExcelSystems AS DECIMAL(5,2)) AS MatchPct
FROM ita.snowserver s
INNER JOIN sts.server sr ON LOWER(sr.ApplicationName) = LOWER(s.System)
WHERE s.System IS NOT NULL
GROUP BY s.System
ORDER BY COUNT(*) DESC

-- Summary Report
SELECT 
    sts_ServerColumn,
    SUM(MatchCount) AS TotalMatches,
    COUNT(DISTINCT ExcelSystemValue) AS UniqueSystemValuesMatched,
    CAST(AVG(MatchPercentage) AS DECIMAL(5,2)) AS AvgMatchPercentage
FROM @Results
GROUP BY sts_ServerColumn
ORDER BY SUM(MatchCount) DESC

-- Detailed matches
PRINT '--- DETAILED MATCHES ---'
SELECT 
    ExcelSystemValue,
    sts_ServerColumn,
    MatchCount,
    MatchPercentage
FROM @Results
WHERE MatchCount > 0
ORDER BY sts_ServerColumn, MatchCount DESC
"@

# Execute the query
try {
    Write-Host "Analyzing System column from Excel against sts.server columns..." -ForegroundColor Cyan
    Write-Host ""
    
    # Adjust connection string if needed
    $connectionString = "Server=localhost;Database=YourDatabase;Integrated Security=true;"
    
    # Try to execute
    $results = Invoke-Sqlcmd -Query $sqlQuery -ConnectionString $connectionString -ErrorAction Stop
    
    Write-Host "ANALYSIS RESULTS:" -ForegroundColor Green
    Write-Host ""
    $results | Format-Table -AutoSize
}
catch {
    Write-Host "Error executing query: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Query to execute:" -ForegroundColor Yellow
    Write-Host $sqlQuery
}
