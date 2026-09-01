-- ANALYSIS: Excel System Column Values vs sts.server Columns
-- Extracted from sts_server_comparison.sql UPDATE statements
-- 501 System entries found with multiple unique values

DECLARE @ExcelSystemValues TABLE (SystemValue NVARCHAR(MAX), RecordCount INT)

-- Unique System values from Excel (sample of top 50 most frequent)
INSERT INTO @ExcelSystemValues (SystemValue, RecordCount) VALUES
(N'Azure Analytics', 95),
(N'DXC Infrastructure', 62),
(N'MMIS', 45),
(N'TCOS - TN Anytime', 28),
(N'InterChange', 22),
(N'MMIS Core', 15),
(N'Doctors Lookup', 12),
(N'SOA Production-DR (RHEL8)', 11),
(N'EOG', 8),
(N'FIS - Formworks 5', 7),
(N'TCOS', 7),
(N'EOG Citrix', 6),
(N'Online Systems', 5),
(N'Tableau', 5),
(N'ATL', 5),
(N'Analytics', 4),
(N'ETL', 4),
(N'IIS Web App / BPM Portals', 4),
(N'ATP', 6),
(N'Code Management', 1),
(N'Autosys', 6),
(N'RightFax', 1),
(N'ALM', 1),
(N'Interchange', 3),
(N'SFTP', 4),
(N'Security', 8),
(N'CrowdStrike (SIEM) for Arcsight', 3),
(N'Qlik Compose', 1),
(N'Access', 1),
(N'Primary Mid-Server for SNOW Discovery', 1),
(N'DA Data Science VM Python', 1),
(N'Secondary Mid-Server for SNOW Discovery', 1),
(N'Analysis Services Gateway', 2),
(N'Reporting', 1),
(N'VM AD Domain Controller', 1),
(N'VM AD Backup Domain Controller', 1),
(N'VM AD Mgt. Server (Restricted)', 1),
(N'Tableau (EOC)', 1),
(N'CrowdStrike', 1)

-- Analysis Query: Check which sts.server columns contain these System values
SELECT 
    'sts.server.System' AS ColumnName,
    COUNT(DISTINCT LOWER(s.System)) AS UniqueMatches,
    COUNT(*) AS TotalMatches,
    CAST(COUNT(*) * 100.0 / (SELECT SUM(RecordCount) FROM @ExcelSystemValues) AS DECIMAL(5,2)) AS MatchPercentage
FROM @ExcelSystemValues e
JOIN sts.server s ON LOWER(s.System) = LOWER(e.SystemValue)

UNION ALL

SELECT 
    'sts.server.ServerDescription',
    COUNT(DISTINCT LOWER(s.ServerDescription)),
    COUNT(*),
    CAST(COUNT(*) * 100.0 / (SELECT SUM(RecordCount) FROM @ExcelSystemValues) AS DECIMAL(5,2))
FROM @ExcelSystemValues e
JOIN sts.server s ON LOWER(s.ServerDescription) = LOWER(e.SystemValue)

UNION ALL

SELECT 
    'sts.server.ServerType',
    COUNT(DISTINCT LOWER(s.ServerType)),
    COUNT(*),
    CAST(COUNT(*) * 100.0 / (SELECT SUM(RecordCount) FROM @ExcelSystemValues) AS DECIMAL(5,2))
FROM @ExcelSystemValues e
JOIN sts.server s ON LOWER(s.ServerType) = LOWER(e.SystemValue)

UNION ALL

SELECT 
    'sts.server.PrimaryFunction',
    COUNT(DISTINCT LOWER(s.PrimaryFunction)),
    COUNT(*),
    CAST(COUNT(*) * 100.0 / (SELECT SUM(RecordCount) FROM @ExcelSystemValues) AS DECIMAL(5,2))
FROM @ExcelSystemValues e
JOIN sts.server s ON LOWER(s.PrimaryFunction) = LOWER(e.SystemValue)

UNION ALL

SELECT 
    'sts.server.ApplicationName',
    COUNT(DISTINCT LOWER(s.ApplicationName)),
    COUNT(*),
    CAST(COUNT(*) * 100.0 / (SELECT SUM(RecordCount) FROM @ExcelSystemValues) AS DECIMAL(5,2))
FROM @ExcelSystemValues e
JOIN sts.server s ON LOWER(s.ApplicationName) = LOWER(e.SystemValue)

UNION ALL

SELECT 
    'sts.server.Environment',
    COUNT(DISTINCT LOWER(s.Environment)),
    COUNT(*),
    CAST(COUNT(*) * 100.0 / (SELECT SUM(RecordCount) FROM @ExcelSystemValues) AS DECIMAL(5,2))
FROM @ExcelSystemValues e
JOIN sts.server s ON LOWER(s.Environment) = LOWER(e.SystemValue)

ORDER BY TotalMatches DESC
