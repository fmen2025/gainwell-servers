# Extract System column values from SQL updates and analyze against sts.server

$systemValues = @"
DXC Infrastructure
MMIS
TCOS - TN Anytime
Azure Analytics
Security
VM AD Domain Controller
VM AD Backup Domain Controller
VM AD Mgt. Server (Restricted)
Tableau (EOC)
RHEL 8 EOC migration dev
IIS (StaffMap) SQL Server
TennCare Call Recording
CrowdStrike
Azure Analytics
CIS CAT Pro security compliance workstation
"@ -split "`n" | Where-Object {$_.Trim()}

Write-Host "Found System column values from Excel:" -ForegroundColor Green
$systemValues | ForEach-Object {Write-Host "  - $_"}
Write-Host ""
Write-Host "Analyzing matches in sts.server table..." -ForegroundColor Cyan
Write-Host ""

$sqlQuery = @"
-- Analysis: Excel System column values in sts.server columns

DECLARE @SystemValues TABLE (SystemValue NVARCHAR(255))
INSERT INTO @SystemValues VALUES 
  (N'DXC Infrastructure'),
  (N'MMIS'),
  (N'TCOS - TN Anytime'),
  (N'Azure Analytics'),
  (N'Security'),
  (N'VM AD Domain Controller'),
  (N'VM AD Backup Domain Controller'),
  (N'VM AD Mgt. Server (Restricted)'),
  (N'Tableau (EOC)'),
  (N'CrowdStrike')

SELECT 'sts.server.System' AS ColumnChecked, COUNT(*) AS Matches
FROM @SystemValues sv
JOIN sts.server s ON LOWER(s.System) = LOWER(sv.SystemValue)
UNION ALL
SELECT 'sts.server.ServerDescription', COUNT(*)
FROM @SystemValues sv
JOIN sts.server s ON LOWER(s.ServerDescription) = LOWER(sv.SystemValue)
UNION ALL
SELECT 'sts.server.ServerType', COUNT(*)
FROM @SystemValues sv
JOIN sts.server s ON LOWER(s.ServerType) = LOWER(sv.SystemValue)
UNION ALL
SELECT 'sts.server.PrimaryFunction', COUNT(*)
FROM @SystemValues sv
JOIN sts.server s ON LOWER(s.PrimaryFunction) = LOWER(sv.SystemValue)
UNION ALL
SELECT 'sts.server.ApplicationName', COUNT(*)
FROM @SystemValues sv
JOIN sts.server s ON LOWER(s.ApplicationName) = LOWER(sv.SystemValue)
ORDER BY Matches DESC
"@

Write-Host "SQL Query to execute:" -ForegroundColor Yellow
Write-Host $sqlQuery
