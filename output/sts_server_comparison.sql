-- Comparison: Excel ServerNames vs sts.server table
-- Generated: 2026-08-26
-- Source Excel: Gainwell Servers_20260820.xlsx (data tab, Name column)
-- Target Table: sts.server
-- Total servers from Excel: 951
-- Function used: ita.fnSnowServer(N'....', 'Servername')

step 1-- Find servers from Excel that EXIST in sts.server, using ita.fnSnowServer function
-- ============================================================================
-- This query uses the ita.fnSnowServer function to look up each Excel server
-- Returns servers from Excel that DO exist in sts.server
--on this 946 servers are existing in sts.server
-- 5 server are null geting from ita.fnSnowServer function
--so the total servers found in sts.server using ita.fnSnowServer function is 946

WITH excel_servers AS (
  SELECT ita.fnSnowServer(N'AH5719006UB007', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'AH57NDCWW01152', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'AH57NDCWW01153', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'AH57NDCWW01165', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'AZE1LMCDEVREF01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'AZE1LMCDEVSYS01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'AZE1WDLDRHAPP01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'AZE1WDLPRDAPP01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'AZE1WDLSTGAPP01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'AZE2LMCDEVSYS02', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'AZE2WDLDRHAPP02', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'AZE2WDLPRDAPP02', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'AZE2WDLSTGAPP02', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'DC5719WA170', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'DC5719WA62', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'DC5719WA63', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'DC5719WA64', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'DC5719WA65', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'DC5719WA66', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'DC5719WA68', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'DC5719WA69', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'DC5719WAP191', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'DC5719WAP31', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'DC5719WWP35', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'DC5719WWP44', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'OV4VC-9.0.0.1181', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ag03ndcui00114', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ag03ndcwt70022', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ag03ndcwt70023', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ag03ndcwt70024', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ag03ndcwt70025', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ag03ndcwt70026', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ag03sdcui00104', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ag03sdcwt00006', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ag03sdcwt70019', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ag03sdcwt70020', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ag03sdcwt70021', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ag03sdcwt70022', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ag03sdcwt70023', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ag03sdcwt70024', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ag03sdcwt70025', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ag03sdcwt70026', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ag03sdcwt80012', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ag03sdcwt80013', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ag03sdcwt80014', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah5719006ub003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah5719006ub005', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00la04001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00la04002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00la04003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00la04004', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00la04005', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00la04006', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00la04007', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00la04008', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00la04009', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00la04010', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00la04011', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00la04012', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00la04013', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00la04014', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00la04015', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00la04016', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00wa04001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00wa04002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00wa04003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00wa04008', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00wa04009', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00wa04010', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00wa04011', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00wa04101', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00wa04105', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00wa04301', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00wa4201', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00wa4202', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00wa4305', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00ww04001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57c00ww04002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00004', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00005', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00008', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00012', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00054', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00061', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00066', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00067', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00068', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00147', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00148', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00149', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00168', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00169', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00170', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00171', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00172', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00173', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00174', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00175', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00176', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla00177', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01040', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01041', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01042', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01043', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01044', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01045', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01046', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01047', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01048', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01049', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01050', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01051', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01052', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01053', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01054', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01055', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01056', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01057', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01058', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01059', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01060', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01061', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01062', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01063', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01064', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01066', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01067', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01068', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01069', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01070', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01071', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01072', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01073', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01074', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01075', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla01076', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla04204', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla04205', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla04206', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla20002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla90013', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla90014', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcla90015', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclb03003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclc00002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclc00007', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclc00008', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclc00010', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclc00011', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclc03005', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclc03006', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclc03007', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclc03008', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclc04003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclc04004', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcli00003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcli00004', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcli00006', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclm00001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclw00096', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclw00097', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclw00098', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclw00099', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclw00111', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclw00112', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclx00002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclx00004', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndclx00005', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcpa00003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcpa00004', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcpa00005', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcpi00005-c', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcpi00006-c', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04004', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04005', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04006', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04007', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04008', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04009', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04010', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04011', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04012', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04013', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04014', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04015', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04016', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04017', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04018', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04019', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04020', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04021', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04022', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04023', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04024', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04025', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04026', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04027', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04028', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04029', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04030', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04031', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04032', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04033', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04034', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04035', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04036', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04037', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04038', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04039', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04040', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04041', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04042', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04043', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04044', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04045', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04046', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04047', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04051', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04052', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04053', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04054', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04055', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04056', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04057', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04058', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04059', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcub04060', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcuh04001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcuh04002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcuh04003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcuh04004', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcuh04005', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00051', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00052', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00053', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00054', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00055', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00056', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00057', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00058', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00095', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00104', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00105', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00106', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00107', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00108', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00109', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00164', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00165', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00166', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00202', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa00203', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa20001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa20003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa20004', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa20005', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa20006', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwa20007', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww00030', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww00034', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww00035', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww00036', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww00040', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww00063', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww00069', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww00119', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww00120', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww01080', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww01084', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww01086', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww01088', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww01090', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww01154', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww01155', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww01156', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww01157', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww01158', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww01166', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcww01167', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwx00001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57ndcwx00002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00061', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00062', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00065', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00066', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00067', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00068', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00069', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00070', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00071', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00072', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00080', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00081', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00082', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00083', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00095', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00096', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00102', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00112', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00113', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00114', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00115', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00127', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00128', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00129', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00130', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00131', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00132', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00133', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00134', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00135', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00136', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00137', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00138', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00139', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00140', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00141', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00142', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00143', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00144', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00145', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00146', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00147', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00148', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00149', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00150', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00151', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00152', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00153', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00154', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00155', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00156', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00157', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00158', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00159', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00160', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00161', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00162', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00163', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00164', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00165', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00166', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00167', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00168', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00169', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00170', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00171', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00172', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00173', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00174', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00175', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00176', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00177', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla00178', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla01050', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla01051', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla01052', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla01053', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla01054', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla01055', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla01056', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla01057', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla01058', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla01059', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla01060', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla01061', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla01062', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla03023', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla03024', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla03025', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla03026', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla04201', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla04202', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla04203', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla04204', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla04205', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla04206', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla04207', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla90013', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla90014', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcla90015', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclb03003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclb03004', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclc00001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclc00002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclc00007', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclc00008', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclc03003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclc03004', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclc03005', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclc03006', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclc03007', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclc03008', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclc03009', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclc03010', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclc04003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclc04004', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcli00002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcli00003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclm00001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclw00060', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclw00061', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclw00096', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclw00097', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclw00103', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclw00107', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclw00108', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclw00109', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclw00110', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclw00111', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclw00112', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclw20002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclx00002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclx00004', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdclx00005', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcpa00003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcpa00004', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcpa00005', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcpa00008', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcpa00009', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcpa00010', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcpi00005-c', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcpi00006-c', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcua00001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcua00002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub00016', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub01018', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub01020', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub01021', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub01027', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub02001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub02001- dup', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04004', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04005', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04006', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04007', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04008', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04009', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04010', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04011', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04012', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04013', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04014', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04015', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04016', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04017', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04018', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04019', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04020', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04021', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04022', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04023', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04024', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04025', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04026', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04027', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04028', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04029', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04030', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04031', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04032', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04033', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04034', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04035', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04036', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04037', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04038', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04039', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04040', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04041', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04043', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04045', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04046', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04048', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04049', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04056', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04057', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04060', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04061', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04062', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04063', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04064', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcub04066', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcuh04001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcuh04002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcuh04003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcuh04004', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcuh04005', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcux01013', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcux01015', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcux01017', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcux01019', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcux01025', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcux01027', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcux01028', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcux01029', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcux01030', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcux01031', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcux01033', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcux01034', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcux01035', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcux01036', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcux01037', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcux01038', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcux01041', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00051', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00056', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00057', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00058', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00066', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00067', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00068', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00069', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00070', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00071', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00072', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00073', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00074', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00075', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00076', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00077', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00078', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00079', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00080', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00081', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00082', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00083', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00084', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00085', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00086', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00087', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00088', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00089', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00090', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00095', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00096', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00097', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00100', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00101', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00102', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00103', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00104', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00105', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00106', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00107', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00108', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00109', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00110', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00111', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00112', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00113', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00114', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00115', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00116', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00117', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00118', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00119', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00120', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00121', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00122', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00123', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00124', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00125', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00126', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00127', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00128', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00129', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00130', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00131', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00132', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00133', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00134', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00135', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00136', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00137', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00138', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00139', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00140', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00141', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00142', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00143', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00144', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00145', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00146', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00147', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00148', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00149', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00150', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00152', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00154', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00155', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00156', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00157', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00158', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00159', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00160', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00161', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00162', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00163', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00164', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00165', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00166', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00167', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00168', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00169', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00170', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00171', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00172', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00173', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00174', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00175', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00176', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00177', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00178', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00179', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00180', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00181', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00182', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00183', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00184', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00185', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00186', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00187', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00188', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00189', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00190', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00191', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00192', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00197', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00198', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00202', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00203', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00204', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa00205', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa01137', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20008', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20009', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20010', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20011', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20012', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20013', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20014', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20015', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20016', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20017', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20018', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20019', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20020', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20021', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20022', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20023', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20024', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20025', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20026', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20027', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwa20028', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwb00002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwb00013', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwb00019', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwb00020', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwb00021', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwb00027', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwb00040', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwb00041', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwb00042', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwb00043', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwc00001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwc00002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwc00003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwc00004', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwf00001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwt00001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwt00002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00029', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00030', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00035', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00036', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00037', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00038', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00040', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00042', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00044', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00045', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00046', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00053', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00054', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00055', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00059', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00060', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00061', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00062', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00063', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00064', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00065', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00066', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00067', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00068', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00069', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00070', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00071', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00072', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00073', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00074', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00075', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00076', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00077', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00078', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00079', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00080', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00081', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00082', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00086', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00087', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00088', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00100', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00101', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00102', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00103', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00104', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00112', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00113', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00114', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00115', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00116', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00117', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00118', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00119', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00120', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00121', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww00122', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01081', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01082', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01083', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01084', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01085', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01086', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01087', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01088', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01089', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01090', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01091', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01150', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01151', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01152', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01153', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01154', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01155', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01156', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01157', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01158', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01165', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01166', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01167', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01168', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01169', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01170', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01171', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01172', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww01173', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww20002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww20003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcww20004', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwx00001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwx00002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwx00003', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwx00004', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwx00007', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'ah57sdcwx00008', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719052wa015', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719052wa016', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719052wa017', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719052waa221', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719052wai102', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719052wb001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719052wb002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719052wb100', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719??', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719cfs60', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719dy249', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719dyt250', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719icdev103', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719pa188', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719pxe32', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719rri67', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719rri68', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719rri69', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719swap99', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719vctr245', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wa107', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wa13', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wa131', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wa171', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wa172', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wa173', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wa174', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wa175', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wa35', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wa75', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wa76', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719waa193', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719waa66', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719waa73', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719waa74', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719waf79', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap10', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap100', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap101', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap102', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap105', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap106', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap108', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap109', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap11', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap115', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap116', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap12', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap125', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap127', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap128', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap129', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap141', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap143', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap144', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap145', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap146', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap147', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap148', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap15', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap150', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap151', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap168h', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap170', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap19', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap190', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap196', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap244', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap34', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap35', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap39', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap42', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap50', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap56', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap57', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap60', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap61', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap65', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap67', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap68', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap69', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap70', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap71', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap72', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap73', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap74', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap75', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap76', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap77', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap95', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap97', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wap98', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wat119', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wb02', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wb62', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wbb161', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wbb162', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wbb164', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wbb165', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wbb166', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wbb30', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wdb122', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wfs10', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wl01313-t', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wss180', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwa110', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwa111', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwa155', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwa156', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwa157', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwp210', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwp211', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwp212', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwp213', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwp214', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwp215', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwp216', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwp224', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwp40', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwp41', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwp43', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwp45', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwp46', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwp47', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wwp48', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'dc5719wws140', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'sgprd47calbas01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'tennacccwdstk', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'tnxixcpctr1', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'tnxixcpctr2', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'tnxixcpmgm01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'tnxixe1warc001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'tnxixe1wccs001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'tnxixe1wd001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'tnxixe1wd002', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'tnxixe1wmgm001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'tnxixe1wqly001', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmdeveueoc01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmdeveusas01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmdeveusas02', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmdeveusas03', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmdeveusas04', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmdeveutab01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeuadmin01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeuans01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeucat01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeudc02', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeudgw01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeudgw02', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeuds01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeueoc01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeueoc02', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeueoc03', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeujump01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeupa02', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeupa03', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeupa04', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeupa05', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeurds01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeurds10', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeusas01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeusas02', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeusas03', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeusas04', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeuspk01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeuspk02', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeuspk03', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeuspk04', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeustw01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmprdeutab01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmsbxeuadmin01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmsbxeuds01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmuateuds01', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'vmuateurds10', 'Servername') AS Servername UNION ALL
  SELECT ita.fnSnowServer(N'win-2ijjqtt29ru', 'Servername') AS Servername
)
SELECT e.Servername FROM excel_servers e
WHERE EXISTS (SELECT 1 FROM sts.server s WHERE LOWER(s.Servername) = LOWER(ita.fnSnowServer(e.Servername, 'Servername')))
ORDER BY e.Servername;

-- step 2-- Map the ita.snowserver mapped columns to their corresponding sts.server columns
-- ============================================================================
-- confirmed = exact column name match; guess = same concept, different name (not data-verified); ? = no candidate found
--
--   ita.snowserver column       sts.server column          status
--   Servername                  ServerName                 confirmed (join key)
--   System                      System                     confirmed
--   IPaddress                   IPAddress                  confirmed
--   location                    Location                   confirmed
--   Operationalstatus           OperationalStatus          confirmed
--   OSversion                   OSversion                  confirmed
--   Environment                 ENV                        guess
--   OS                          OperatingSystem            guess
--   Purpose                     ServerDescription          guess
--   Usedfor                     PrimaryFunction            guess
--   Installstatus                Status                     guess
--   Hostadmin                   OSAdmin                    guess
--   Primaryapplicationadmin     SAOwner                    guess
--   ManagedBy                   TechnicalOwner             guess
--   Ownedby                     BusinessOwner              guess
--   Sysclassname                ServerType                 guess
--   Lastdiscovered              DeviceWasLastPingedAt      guess
--   Sysupdatedon                ModDate                    guess
--   Sysupdatedby                ModUser                    guess
--   FirstDiscovered             ADDDate                    guess
--   Vendor                      ?                          no candidate
--   Supportgroup                ?                          no candidate
--   Supportedby                 ?                          no candidate (OSAdmin already used for Hostadmin)
--   Hostadminprimary            ?                          no candidate (OSAdmin already used for Hostadmin)
--   Managedbygroup              ?                          no candidate
--   Classification              ?                          no candidate
--   Maintenanceschedule         ?                           no candidate

-- step 2-- Data-overlap verification of the above candidates (read-only, joined on Servername, case-insensitive)
-- ============================================================================
-- joined = rows matched between the two tables; both_null/null_ita/null_sts = which side is NULL;
-- match/differ = non-NULL value comparison count; match% = match / (match + differ)
--
--   ita.snowserver           sts.server              status      joined both_null null_ita null_sts  match differ match% verdict
--   IPaddress                IPAddress               confirmed     3161       34       29        0    3088     10   100%  verified match
--   OS                       OperatingSystem         guess         3161       88       37        0    2939     97    97%  verified match (promote to confirmed)
--   OSversion                OSversion               confirmed     3161      248       24        0    2856     33    99%  verified match
--   Operationalstatus        OperationalStatus       confirmed     3161        0        0        0    3053    108    97%  verified match
--   Ownedby                  BusinessOwner           guess         3161      112     3048        0       1      0   100%  only 1 comparable row - not conclusive
--   System                   System                  confirmed     3161      186     2975        0       0      0    n/a  can't verify - ita.snowserver.System is 100% NULL
--   location                 Location                confirmed     3161       13        0     3148       0      0    n/a  can't verify - sts.server.Location is ~99.6% NULL
--   Purpose                  ServerDescription       guess         3161        0     3161        0       0      0    n/a  reject - sts.server.ServerDescription is 100% NULL
--   Installstatus            Status                  guess         3161        0        0     3161       0      0    n/a  reject - sts.server.Status is 100% NULL
--   Primaryapplicationadmin  SAOwner                 guess         3161     1568     1526       61       2      4    33%  reject - tiny overlap
--   ManagedBy                TechnicalOwner          guess         3161       62      120      139     110   2730     4%   reject
--   Environment              ENV                     guess         3161        0     1030        0       0   2131     0%   reject
--   Usedfor                  PrimaryFunction         guess         3161        0        0     2047       0   1114     0%   reject
--   Hostadmin                OSAdmin                 guess         3161      295        8     2290       0    568     0%   reject
--   Sysclassname             ServerType              guess         3161        0        0      802       2   2357     0%   reject
--   Lastdiscovered           DeviceWasLastPingedAt   guess         3161        0      282        0       0   2879     0%   reject
--   Sysupdatedon             ModDate                 guess         3161        0        0        0       0   3161     0%   reject
--   Sysupdatedby             ModUser                 guess         3161        0        0        0       0   3161     0%   reject
--   FirstDiscovered          ADDDate                 guess         3161        0      243        0       0   2918     0%   reject
--
-- Net result: 4 columns verified by data (IPaddress, OSversion, Operationalstatus, OS->OperatingSystem);
-- 2 confirmed-by-name columns (System, location) can't be verified either way (one side is almost all NULL);
-- 13 of 14 "guess" mappings are rejected - matching column names only, not matching data.

-- step 3-- Cross-check the user's existing stored procedure mapping (alias t -> tr) against step 2/2b findings
-- ============================================================================
-- tr = sts.server (target); t = ita.snowserver (confirmed by user 2026-08-28) - so
-- ChangeorderNumber/company/ram/dsnDomain/osdomain/Process are real ita.snowserver columns
-- outside the 26-column subset step 2 tested (that subset was scoped to mapping.md's
-- Excel-driven columns only, not every column ita.snowserver actually has).
-- Pattern used in the SP: tr.col = COALESCE(LTRIM(RTRIM(t.col)), tr.col) - source overwrites the
-- target whenever it is non-NULL, otherwise the existing tr value is kept (not a NULL-only fill).
--
--   SP: t.column -> tr.column      cross-check vs step 2/2b                          verdict
--   ipaddress -> IPAddress         matches our tested pair (100% match)              confirmed
--   operationalstatus -> operationalstatus  matches our tested pair (97% match)      confirmed
--   os -> Operatingsystem          matches our tested pair (97% match)               confirmed
--   OSversion -> OSversion         matches our tested pair (99% match)               confirmed
--   location -> DataCenter         NOT Location - smarter than our guess since       makes sense
--                                  sts.server.Location is ~99.6% NULL
--   Process -> system              NOT System - consistent with ita.snowserver.System   makes sense
--                                  being 100% NULL; sources from a populated column instead
--   environment -> primarysubsystem  opposite of our naive name-match guess           not yet verified
--   classification -> ENV            (we'd have guessed environment->ENV) - both t and    with real data
--                                     tr are now known/reachable, so this can be tested
--   managedby -> Businessowner     untested pairing; we only tested ManagedBy->        not yet verified
--   supportedby -> TechnicalOwner  TechnicalOwner and rejected it (4% match)           with real data
--   sysclassname -> ServerDescription  sts.server.ServerDescription is 100% NULL, so    not yet verified
--                                      this is a safe fill but accuracy is unproven     with real data
--   ChangeorderNumber, company->Agency, ram->memory, dsnDomain->LocationExtension,
--   osdomain->AssetOwner            real ita.snowserver columns, just outside the        not yet verified
--                                   26-column subset step 2 originally tested            with real data
--
-- Next: re-run the step 2b-style data-overlap query for these remaining pairs (all within
-- ita.snowserver/sts.server, both reachable) to confirm or reject them the same way.

--step 4 :- create temp tables to store  results of ita.snowserver , sts.server
--using select * into tp store all server and snow table below separately
  SELECT * INTO #TempSnowServer FROM ita.snowserver;
  SELECT * INTO #TempServer FROM sts.server;

 --step 5:- the update script that you creted on the compare_servers_consolidated.sql 
 --`  RECOMMENDED UPDATEs for NULL fields (CONSOLIDATED - 1 UPDATE per server)
-- Total consolidated updates: 945` update script will follow based on the comparison results stored in the temp tables.
-- ============================================================================
-- Direction: Excel -> ita.snowserver. Live against the temp table (safe - #TempSnowServer
-- is a session-scoped copy from step 4, not the real ita.snowserver):
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'Azure Domain Controller', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'(unknown)' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Interchange', [Purpose] = N'Created Temp for Decom.', [Supportgroup] = N'Unix Admin' WHERE Servername = N'AH5719006UB007' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS', [Purpose] = N'Web Server', [Supportgroup] = N'TCOS SW ADMIN' WHERE Servername = N'AH57NDCWW01152' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS', [Purpose] = N'Web Server', [Supportgroup] = N'TCOS SW ADMIN' WHERE Servername = N'AH57NDCWW01153' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Online Systems', [Purpose] = N'OS Web / App', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'AH57NDCWW01165' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS Core', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'AZE1LMCDEVREF01' AND ([System] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS Core', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'AZE1LMCDEVSYS01' AND ([System] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Doctors Lookup', [Purpose] = N'Doctor Lookup Dr app Server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'AZE1WDLDRHAPP01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Doctors Lookup', [Purpose] = N'Doctor Lookup Prod app Server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'AZE1WDLPRDAPP01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Doctors Lookup', [Purpose] = N'Doctor Lookup Test app Server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'AZE1WDLSTGAPP01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS Core', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'AZE2LMCDEVSYS02' AND ([System] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Doctors Lookup', [Purpose] = N'Doctor Lookup Dr app Server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'AZE2WDLDRHAPP02' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Doctors Lookup', [Purpose] = N'Doctor Lookup Prod app Server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'AZE2WDLPRDAPP02' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Doctors Lookup', [Purpose] = N'Doctor Lookup Test app Server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'AZE2WDLSTGAPP02' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'RightFax', [Purpose] = N'RightFax', [Supportgroup] = N'Unix Admin' WHERE Servername = N'DC5719WA170' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FIS - Formworks 5', [Purpose] = N'FIS - Formworks 5', [Supportgroup] = N'Unix Admin' WHERE Servername = N'DC5719WA62' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FIS - Formworks 5', [Purpose] = N'FIS - Formworks 5', [Supportgroup] = N'Unix Admin' WHERE Servername = N'DC5719WA63' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FIS - Formworks 5', [Purpose] = N'FIS - Formworks 5', [Supportgroup] = N'Unix Admin' WHERE Servername = N'DC5719WA64' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FIS - Formworks 5', [Purpose] = N'FIS - Formworks 5', [Supportgroup] = N'Unix Admin' WHERE Servername = N'DC5719WA65' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FIS - Formworks 5', [Purpose] = N'FIS - Formworks 5', [Supportgroup] = N'Unix Admin' WHERE Servername = N'DC5719WA66' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FIS - Formworks 5', [Purpose] = N'FIS - Formworks 5', [Supportgroup] = N'Unix Admin' WHERE Servername = N'DC5719WA68' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FIS - Formworks 5', [Purpose] = N'FIS - Formworks 5', [Supportgroup] = N'Unix Admin' WHERE Servername = N'DC5719WA69' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM', [Purpose] = N'ALM Test Data Masking', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'DC5719WAP191' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'vSphere web interface', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'DC5719WAP31' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS - TN Anytime', [Purpose] = N'TCOS - TN Anytime', [Supportgroup] = N'Unix Admin' WHERE Servername = N'DC5719WWP35' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS - TN Anytime', [Purpose] = N'TCOS - TN Anytime', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'DC5719WWP44' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'HPE OneView for VMware vCenter', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'OV4VC-9.0.0.1181' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SFTP', [Purpose] = N'LDOM/Zone', [Supportgroup] = N'Legacy SFTP SW ADMIN' WHERE Servername = N'ag03ndcui00114' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Purpose] = N'TennCare Interchange XenApp', [Supportgroup] = N'INTERCHANGE SW ADMIN' WHERE Servername = N'ag03ndcwt70022' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Purpose] = N'TennCare Interchange XenApp', [Supportgroup] = N'INTERCHANGE SW ADMIN' WHERE Servername = N'ag03ndcwt70023' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Purpose] = N'TennCare Interchange XenApp', [Supportgroup] = N'INTERCHANGE SW ADMIN' WHERE Servername = N'ag03ndcwt70024' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Purpose] = N'TennCare Interchange XenApp', [Supportgroup] = N'INTERCHANGE SW ADMIN' WHERE Servername = N'ag03ndcwt70025' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Purpose] = N'TennCare Interchange XenApp', [Supportgroup] = N'INTERCHANGE SW ADMIN' WHERE Servername = N'ag03ndcwt70026' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SFTP', [Purpose] = N'LDOM/Zone Per INC0561804 - This server was retired since it was retired by STS in 2021. Retired manually.', [Supportgroup] = N'Legacy SFTP SW ADMIN' WHERE Servername = N'ag03sdcui00104' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Purpose] = N'Citrix XenApp Server', [Supportgroup] = N'TennCare Technical Service Desk' WHERE Servername = N'ag03sdcwt00006' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'EOG', [Purpose] = N'Citrix Virtual Apps', [Supportgroup] = N'EOG SW ADMIN' WHERE Servername = N'ag03sdcwt70019' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'EOG', [Purpose] = N'Citrix Virtual Apps', [Supportgroup] = N'EOG SW ADMIN' WHERE Servername = N'ag03sdcwt70020' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'EOG', [Purpose] = N'Citrix Virtual Apps', [Supportgroup] = N'EOG SW ADMIN' WHERE Servername = N'ag03sdcwt70021' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Purpose] = N'TennCare Interchange XenApp', [Supportgroup] = N'INTERCHANGE SW ADMIN' WHERE Servername = N'ag03sdcwt70022' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Purpose] = N'TennCare Interchange XenApp', [Supportgroup] = N'INTERCHANGE SW ADMIN' WHERE Servername = N'ag03sdcwt70023' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Purpose] = N'TennCare Interchange XenApp', [Supportgroup] = N'INTERCHANGE SW ADMIN' WHERE Servername = N'ag03sdcwt70024' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Purpose] = N'TennCare Interchange XenApp', [Supportgroup] = N'INTERCHANGE SW ADMIN' WHERE Servername = N'ag03sdcwt70025' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Purpose] = N'TennCare Interchange XenApp', [Supportgroup] = N'INTERCHANGE SW ADMIN' WHERE Servername = N'ag03sdcwt70026' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'EOG Citrix', [Purpose] = N'Citrix Virtual Apps', [Supportgroup] = N'MMIS Support Backend Apps EOG' WHERE Servername = N'ag03sdcwt80012' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'EOG Citrix', [Purpose] = N'Citrix Virtual Apps', [Supportgroup] = N'MMIS Support Backend Apps EOG' WHERE Servername = N'ag03sdcwt80013' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'EOG Citrix', [Purpose] = N'Citrix Virtual Apps', [Supportgroup] = N'MMIS Support Backend Apps EOG' WHERE Servername = N'ag03sdcwt80014' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Interchange', [Purpose] = N'Oracle DR dbs - tnhisp1 and tnmisp1', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah5719006ub003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Purpose] = N'Oracle Financial', [Supportgroup] = N'UNIX OS ADMIN' WHERE Servername = N'ah5719006ub005' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'CrowdStrike (SIEM) for Arcsight', [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'ah57c00la04001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'SAS UI', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'ah57c00la04002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'SAS Master', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'ah57c00la04003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'SAS Worker', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'ah57c00la04004' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'SAS Worker', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'ah57c00la04005' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'SAS UI', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'ah57c00la04006' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'SAS Master', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'ah57c00la04007' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'SAS Worker', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'ah57c00la04008' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'SAS Worker', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'ah57c00la04009' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CrowdStrike (SIEM) for Arcsight', [Purpose] = N'Security', [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'ah57c00la04010' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'EOC (Corticon)', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'ah57c00la04011' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'EOC (Corticon)', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'ah57c00la04012' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'EOC (Corticon)', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'ah57c00la04013' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'Splunk Heavy Forwarder', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'ah57c00la04014' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CrowdStrike (SIEM) for Arcsight', [Purpose] = N'Security', [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'ah57c00la04015' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Security', [Purpose] = N'Security', [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'ah57c00la04016' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Tableau', [Purpose] = N'Azure Analytics', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'ah57c00wa04001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ETL', [Purpose] = N'Azure Analytics', [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'ah57c00wa04002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Analytics', [Purpose] = N'Azure Analytics', [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'ah57c00wa04003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Tableau', [Purpose] = N'Azure Analytics', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'ah57c00wa04008' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Tableau', [Purpose] = N'Azure Analytics', [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'ah57c00wa04009' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Qlik Compose', [Purpose] = N'ETL', [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'ah57c00wa04010' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Access', [Purpose] = N'Azure Analytics', [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'ah57c00wa04011' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Primary Mid-Server for SNOW Discovery', [Purpose] = N'Azure Analytics', [Supportgroup] = N'ServiceNow Tier 2 Support' WHERE Servername = N'ah57c00wa04101' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DA Data Science VM Python', [Purpose] = N'Azure Analytics', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'ah57c00wa04105' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Secondary Mid-Server for SNOW Discovery', [Purpose] = N'Azure Analytics', [Supportgroup] = N'ServiceNow Tier 2 Support' WHERE Servername = N'ah57c00wa04301' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Analysis Services Gateway', [Purpose] = N'Azure Analytics', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'ah57c00wa4201' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Reporting', [Purpose] = N'Azure Analytics', [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'ah57c00wa4202' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Analysis Services Gateway', [Purpose] = N'Azure Analytics', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'ah57c00wa4305' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'IIS Web App / BPM Portals', [Purpose] = N'Azure Analytics', [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'ah57c00ww04001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'IIS Web App / BPM Portals', [Purpose] = N'Azure Analytics', [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'ah57c00ww04002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ATP', [Purpose] = N'IIB App Server - Uses IPTables - OS patching can restart without app admin assistance.', [Supportgroup] = N'ATP SW ADMIN' WHERE Servername = N'ah57ndcla00004' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ATP', [Purpose] = N'IIB App Server - Uses IPTables - OS patching can restart without app admin assistance.', [Supportgroup] = N'ATP SW ADMIN' WHERE Servername = N'ah57ndcla00005' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ATP', [Purpose] = N'IIB App Server - OS patching can restart without app admin assistance.', [Supportgroup] = N'ATP SW ADMIN' WHERE Servername = N'ah57ndcla00008' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ATP', [Purpose] = N'IIB App Server - OS patching can restart without app admin assistance.', [Supportgroup] = N'ATP SW ADMIN' WHERE Servername = N'ah57ndcla00012' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ATP', [Purpose] = N'Oracle Database Firewall', [Supportgroup] = N'OAV SW ADMIN' WHERE Servername = N'ah57ndcla00054' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Code Management', [Purpose] = N'Code Management - OS patching can restart without app admin assistance.', [Supportgroup] = N'CODE MANAGEMENT SW ADMIN' WHERE Servername = N'ah57ndcla00061' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Autosys', [Purpose] = N'Autosys App', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57ndcla00066' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Autosys', [Purpose] = N'Autosys App', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57ndcla00067' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Autosys', [Purpose] = N'Autosys App', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57ndcla00068' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Autosys', [Purpose] = N'Autosys App', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57ndcla00147' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Autosys', [Purpose] = N'Autosys App', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57ndcla00148' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Autosys', [Purpose] = N'Autosys App', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57ndcla00149' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'IBM ACE', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla00168' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'IBM ACE', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla00169' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'IBM WAS', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla00170' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'IBM WAS', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla00171' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'IBM BAW', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla00172' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'IBM BAW', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla00173' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'Progress Corticon', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla00174' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'Progress Corticon', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla00175' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'IBM WSRR', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla00176' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'IBM WSRR', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla00177' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Dev', [Purpose] = N'SOA Application Server (WAS)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01040' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Dev', [Purpose] = N'SOA Application Server (BPM)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01041' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Dev', [Purpose] = N'SOA Application Server (WAS)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01042' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Dev', [Purpose] = N'SOA Application Server (IIB)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01043' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Dev', [Purpose] = N'SOA Application Server (WSRR)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01044' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Test', [Purpose] = N'SOA Application Server (BPM)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01045' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Test', [Purpose] = N'SOA Application Server (WAS)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01046' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Test', [Purpose] = N'SOA Application Server (IIB)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01047' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Test', [Purpose] = N'SOA Application Server (WAS)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01048' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Test', [Purpose] = N'SOA Application Server (WSRR)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01049' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (IIB)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01050' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (IIB)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01051' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (WAS)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01052' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (WAS)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01053' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (BPM)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01054' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (BPM)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01055' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (WAS)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01056' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (WAS)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01057' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (WSRR)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01058' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (WSRR)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01059' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA UAT', [Purpose] = N'SOA Application Server (BPM)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01060' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA UAT', [Purpose] = N'SOA Application Server (IODM)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01061' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA UAT', [Purpose] = N'SOA Application Server (IIB)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01062' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA UAT', [Purpose] = N'SOA Application Server (WAS)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01063' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA UAT', [Purpose] = N'SOA Application Server (WSRR)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01064' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage', [Purpose] = N'SOA Application Server (IIB)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01066' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage', [Purpose] = N'SOA Application Server (IIB)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01067' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage', [Purpose] = N'SOA Application Server (IIB)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01068' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage', [Purpose] = N'SOA Application Server (IIB)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01069' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage', [Purpose] = N'SOA Application Server (IIB)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01070' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage', [Purpose] = N'SOA Application Server (IIB)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01071' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage', [Purpose] = N'SOA Application Server (IIB)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01072' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage', [Purpose] = N'SOA Application Server (IIB)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01073' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage', [Purpose] = N'SOA Application Server (IIB)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01074' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage', [Purpose] = N'SOA Application Server (IIB)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01075' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA POC Rebuild', [Purpose] = N'SOA Application Server (WAS)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcla01076' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS Dynatrace', [Purpose] = N'MMIS Dynatrace Cluster Active gate', [Supportgroup] = N'MMIS DynaTrace SW ADMIN' WHERE Servername = N'ah57ndcla04204' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS Dynatrace', [Purpose] = N'MMIS Dynatrace Cluster Active gate', [Supportgroup] = N'MMIS DynaTrace SW ADMIN' WHERE Servername = N'ah57ndcla04205' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS Dynatrace', [Purpose] = N'MMIS Dynatrace Environment Active Gate', [Supportgroup] = N'MMIS DynaTrace SW ADMIN' WHERE Servername = N'ah57ndcla04206' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'AGM', [Purpose] = N'AGM Application Server - OS patching can restart without app admin assistance', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57ndcla20002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [Purpose] = N'NetBackup' WHERE Servername = N'ah57ndcla90013' AND ([Purpose] IS NULL);
UPDATE #TempSnowServer SET [Purpose] = N'NetBackup' WHERE Servername = N'ah57ndcla90014' AND ([Purpose] IS NULL);
UPDATE #TempSnowServer SET [Purpose] = N'NetBackup' WHERE Servername = N'ah57ndcla90015' AND ([Purpose] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Edifecs', [Purpose] = N'Oracle DB', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndclb03003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ATP', [Purpose] = N'Oracle RAC Node - Uses IPTables', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57ndclc00002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Shared RAC', [Purpose] = N'Oracle RAC Node', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57ndclc00007' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Shared RAC', [Purpose] = N'Oracle RAC Node', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57ndclc00008' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Shared RAC', [Purpose] = N'Oracle RAC Node', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57ndclc00010' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Shared RAC', [Purpose] = N'Oracle RAC Node', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57ndclc00011' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle RAC - Sandbox', [Purpose] = N'Oracle Database - RAC - Sandbox DR Node 1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndclc03005' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle RAC - Sandbox', [Purpose] = N'Oracle Database - RAC - Sandbox DR Node 2', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndclc03006' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle RAC', [Purpose] = N'Oracle Database - RAC - DR Node 1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndclc03007' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle RAC', [Purpose] = N'Oracle Database - RAC - DR Node 2', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndclc03008' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle RAC', [Purpose] = N'Oracle Database - RAC - DR Node 1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndclc04003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle RAC', [Purpose] = N'Oracle Database - RAC - DR Node 2', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndclc04004' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA', [Purpose] = N'Datapower', [Supportgroup] = N'ENTERPRISE SOA SW ADMIN' WHERE Servername = N'ah57ndcli00003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA DEV', [Purpose] = N'Datapower Proxy', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcli00004' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA', [Purpose] = N'Datapower', [Supportgroup] = N'ENTERPRISE SOA SW ADMIN' WHERE Servername = N'ah57ndcli00006' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Syslog', [Purpose] = N'Syslog - OS patching can restart without app admin assistance', [Supportgroup] = N'ATP SW ADMIN' WHERE Servername = N'ah57ndclm00001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Web Server', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndclw00096' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Web Server', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndclw00097' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage', [Purpose] = N'SOA Web Server', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndclw00098' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage', [Purpose] = N'SOA Web Server', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndclw00099' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'TMED Web Server', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndclw00111' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'TMED Web Server', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndclw00112' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'OEM', [Purpose] = N'Oracle Enterprise Manager', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57ndclx00002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle Enterprise Manager (OEM)', [Purpose] = N'Oracle Enterprise Manager - OS patching can restart without app admin assistance.', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndclx00004' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle Enterprise Manager (OEM)', [Purpose] = N'Oracle Enterprise Manager - OS patching can restart without app admin assistance.', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndclx00005' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'PAM', [Purpose] = N'PAM Appliance', [Supportgroup] = N'PAM SW Admin' WHERE Servername = N'ah57ndcpa00003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'PAM', [Purpose] = N'PAM Appliance', [Supportgroup] = N'PAM SW Admin' WHERE Servername = N'ah57ndcpa00004' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'PAM', [Purpose] = N'PAM Debian appliance', [Supportgroup] = N'UNIX OS ADMIN' WHERE Servername = N'ah57ndcpa00005' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage', [Purpose] = N'Datapower', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcpi00005-c' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage', [Purpose] = N'Datapower', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57ndcpi00006-c' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS -  OEMN', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57ndcub04002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone 8:31 AM 4/1/2025 Retired per Chad Williams - Anthony Bronsdon', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04004' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57ndcub04005' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57ndcub04006' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SIEBEL', [Purpose] = N'Oracle Database - Siebel - Non-Prod - SBLSBX20, SBLDEV20, SBLUAT20, SBLSTG20', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04007' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T-7 Database Host', [Purpose] = N'Oracle Database - CMOD/FileNet - RMAN restore CMOD/FileNet databases in case of DR', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04008' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - T20HISM1, T20MISM1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04009' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - T21HISM1, T21MISM1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04010' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - T22HISM1, T22MISM1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04011' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - T23HISM1, T23MARM1,  T23MISM1, T23SURM1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04012' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - T24HISM1, T24MARM1,  T24MISM1, T24SURM1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04013' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - T20HISA1, T20MISA1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04014' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - T21HISA1, T21MISA1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04015' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - T22HISM1, T22MARM1,  T22MISM1, T22SURM1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04016' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - T23HISA1, T23MISA1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04017' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone 8:31 AM 4/1/2025 Retired per Chad Williams - Anthony Bronsdon', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04018' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T-7 Database Host', [Purpose] = N'Oracle Database - Siebel - Non-Prod - SBLSBX, SBLDEV17, SBLUAT17', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04019' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57ndcub04020' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle database - MMIS - tnhisa1 tnmara1 tnmisa1 tnsuma1 tnsura1', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57ndcub04021' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database - MMIS - tnhism1 tnmarm1 tnmism1 tnsurm1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04022' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57ndcub04023' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone 8:31 AM 4/1/2025 Retired per Chad Williams - Anthony Bronsdon', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57ndcub04024' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone  - IP in test range.', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57ndcub04025' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'NFS Share', [Supportgroup] = N'MMIS Support INF UNIX', [Managedbygroup] = N'TENNCARE Enterprise Services', [Maintenanceschedule] = N'Quarterly Patching â€“ 2nd Sunday Feb, May, Aug, Nov' WHERE Servername = N'ah57ndcub04026' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL OR [Managedbygroup] IS NULL OR [Maintenanceschedule] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - T30HISA1, T30MARA1, T30MISA1, T30SURA1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04027' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - T25HISM1, T25MISM1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04028' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - T26HISM1, T26MISM1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04029' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - T27HISM1, T27MISM1, T27MARM1, T27SURM1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04030' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - TNHIST2, TNMIST2, TNMART2, TNSURT2', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04031' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - T24HISA1, T24MISA1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04032' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - T25HISA1, T25MISA1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04033' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - T26HISA1, T26MISA1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04034' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - T27HISA1, T27MISA1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04035' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - T28HISA1, T28MARA1, T28MISA1, T28SURA1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04036' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - T29HISA1, T29MARA1, T29MISA1, T29SURA1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04037' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - TNHISA3, TNMARA3, TNMISA3, TNSURA3', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04038' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - TNHISM2, TNMISM2', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04039' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57ndcub04040' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SFTP', [Purpose] = N'LDOM/Zone', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57ndcub04041' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SFTP', [Purpose] = N'LDOM/Zone', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57ndcub04042' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SFTP', [Purpose] = N'LDOM/Zone', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57ndcub04043' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SFTP', [Purpose] = N'LDOM/Zone', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57ndcub04044' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SFTP', [Purpose] = N'LDOM/Zone', [Supportgroup] = N'Legacy SFTP SW ADMIN' WHERE Servername = N'ah57ndcub04045' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SFTP', [Purpose] = N'LDOM/Zone', [Supportgroup] = N'Legacy SFTP SW ADMIN' WHERE Servername = N'ah57ndcub04046' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SFTP', [Purpose] = N'LDOM/Zone', [Supportgroup] = N'Legacy SFTP SW ADMIN' WHERE Servername = N'ah57ndcub04047' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T-7 Database Host', [Purpose] = N'Oracle Database Zone - ALM/DataStage - Non-Prod - TNALMM1, TNDSGA32, TNDSGU32, TNDSTGA31, TNDSTGT3, TNDSTGU31', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04051' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57ndcub04052' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Forgerock Dev 8:31 AM 4/1/2025 Retired per Chad Williams - Anthony Bronsdon', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57ndcub04053' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T-7 Database Host', [Purpose] = N'Oracle Database Zone - CMOD/FileNet - Non-Prod - ARSDV105, ARSTS105, ARSSB105, IBMCNVNA, IBMCNVND, IBMCNVNS, IBMFNVNA, IBMFNVND, IBMFNVNS, IDBP55NA, IDBP55ND, IDBP55NS', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04054' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Forgerock LDAP', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57ndcub04055' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T-8 Database Host', [Purpose] = N'Interchange DR db - MMIS - TNHISDR, TNMISDR, TNDOCP', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04056' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T-8 Database Host', [Purpose] = N'Interchange Data Warehouse (DR) - MMIS - TNMARDR, TNSURDR', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04057' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04058' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - HIS21DE, MIS21DE', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04059' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T8 Database Host', [Purpose] = N'Oracle Database - SAS - DR - SASDR', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57ndcub04060' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Host', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57ndcuh04001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Host', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57ndcuh04002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SFTP', [Purpose] = N'LDOM/Zone Host', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57ndcuh04003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SFTP', [Purpose] = N'LDOM/Zone Host', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57ndcuh04004' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T8 Database Host', [Purpose] = N'T8 Oracle Host', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57ndcuh04005' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Code Management', [Purpose] = N'Code Management - OS patching can restart without app admin assistance.', [Supportgroup] = N'CODE MANAGEMENT SW ADMIN' WHERE Servername = N'ah57ndcwa00051' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel 2016', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57ndcwa00052' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel 2016', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57ndcwa00053' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel 2016', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57ndcwa00054' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel 2016', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57ndcwa00055' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel 2016 App', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57ndcwa00056' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel 2016 App', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57ndcwa00057' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel 2016 App', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57ndcwa00058' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'PAM', [Purpose] = N'CA Authentication', [Supportgroup] = N'PAM SW Admin' WHERE Servername = N'ah57ndcwa00095' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Web/App', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57ndcwa00104' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Web/App', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57ndcwa00105' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Application', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57ndcwa00106' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Application', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57ndcwa00107' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Application', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57ndcwa00108' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Reports/App', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57ndcwa00109' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer App 1', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57ndcwa00164' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer App 2', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57ndcwa00165' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer App 3', [Supportgroup] = N'APPEON SW ADMIN' WHERE Servername = N'ah57ndcwa00166' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer App 1', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57ndcwa00202' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer App 2', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57ndcwa00203' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'AGM', [Purpose] = N'AGM Synchronizer', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57ndcwa20001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM', [Purpose] = N'ALM Synchronizer', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57ndcwa20003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM', [Purpose] = N'ALM Node', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57ndcwa20004' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM', [Purpose] = N'Lab Manager', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57ndcwa20005' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM', [Purpose] = N'Host 1', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57ndcwa20006' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM', [Purpose] = N'Host 2', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57ndcwa20007' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Datastage', [Purpose] = N'Web/App Server', [Supportgroup] = N'DATASTAGE SW ADMIN' WHERE Servername = N'ah57ndcww00030' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel 2016 Web Server', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57ndcww00034' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel 2016 Web Server', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57ndcww00035' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel 2016 Web Server', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57ndcww00036' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'Interchange Web Server', [Supportgroup] = N'APPEON SW ADMIN' WHERE Servername = N'ah57ndcww00040' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Datastage', [Purpose] = N'App/Web Server', [Supportgroup] = N'MMIS Support Data Warehouse DataStage' WHERE Servername = N'ah57ndcww00063' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer Web', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57ndcww00069' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer Web', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57ndcww00119' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer Web', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57ndcww00120' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'EPLS', [Purpose] = N'EPLS Test Server', [Supportgroup] = N'EPLS SW ADMIN' WHERE Servername = N'ah57ndcww01080' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SURS - Surveillance and Utilization Revi', [Purpose] = N'SURS Web / App', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57ndcww01084' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TPM - Trading Partner Management', [Purpose] = N'EDI Trading Partner Web Server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57ndcww01086' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'HIE - Health Information Exchange', [Purpose] = N'HIE Web Server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57ndcww01088' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TMSIS - Transformed Medicaid Statistical', [Purpose] = N'TMSIS Web', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57ndcww01090' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS', [Purpose] = N'Web Server', [Supportgroup] = N'TCOS SW ADMIN' WHERE Servername = N'ah57ndcww01154' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Visual Application', [Purpose] = N'EDI Trading Partner Web Server', [Supportgroup] = N'VISUAL APP SW ADMIN' WHERE Servername = N'ah57ndcww01155' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Visual Application', [Purpose] = N'SURS Web Server', [Supportgroup] = N'VISUAL APP SW ADMIN' WHERE Servername = N'ah57ndcww01156' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Visual Application', [Purpose] = N'HIE Web Server', [Supportgroup] = N'VISUAL APP SW ADMIN' WHERE Servername = N'ah57ndcww01157' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Visual Application', [Purpose] = N'TMSIS_CMS Web Server', [Supportgroup] = N'VISUAL APP SW ADMIN' WHERE Servername = N'ah57ndcww01158' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Online Systems', [Purpose] = N'OS Web / App', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57ndcww01166' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Online Systems', [Purpose] = N'OS Web / App', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57ndcww01167' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Decision Analytics', [Purpose] = N'DR for Prod Cluster 1', [Supportgroup] = N'MMIS Support Decision Analytics Dashboard' WHERE Servername = N'ah57ndcwx00001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Decision Analytics', [Purpose] = N'DR for Prod Cluster 2', [Supportgroup] = N'MMIS Support Decision Analytics Dashboard' WHERE Servername = N'ah57ndcwx00002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ATP', [Purpose] = N'IIB App Server  - OS patching can restart without app admin assistance.', [Supportgroup] = N'ATP SW ADMIN' WHERE Servername = N'ah57sdcla00002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ATP', [Purpose] = N'IIB App Server  - OS patching can restart without app admin assistance.', [Supportgroup] = N'ATP SW ADMIN' WHERE Servername = N'ah57sdcla00003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Code Management', [Purpose] = N'Code Management Build Server - OS patching can restart without app admin assistance.', [Supportgroup] = N'MMIS Support Backend Apps DevOps' WHERE Servername = N'ah57sdcla00061' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Code Management', [Purpose] = N'Code Management Repository Server - OS patching can restart without app admin assistance.', [Supportgroup] = N'MMIS Support Backend Apps DevOps' WHERE Servername = N'ah57sdcla00062' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ATP', [Purpose] = N'Oracle Audit Vault', [Supportgroup] = N'OAV SW ADMIN' WHERE Servername = N'ah57sdcla00065' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Autosys', [Purpose] = N'Autosys App', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcla00066' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Autosys', [Purpose] = N'Autosys App', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcla00067' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Autosys', [Purpose] = N'Autosys App', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcla00068' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Autosys', [Purpose] = N'Autosys App', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcla00069' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Autosys', [Purpose] = N'Autosys App', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcla00070' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Autosys', [Purpose] = N'Autosys App', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcla00071' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'IBM License Metric Tool (IBM LMT)', [Purpose] = N'IBM Licensing Server', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00072' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TMSIS Common Solutions', [Purpose] = N'App Server', [Supportgroup] = N'MMIS Support Data Warehouse TMSIS' WHERE Servername = N'ah57sdcla00080' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TMSIS Common Solutions', [Purpose] = N'App Server', [Supportgroup] = N'MMIS Support Data Warehouse TMSIS' WHERE Servername = N'ah57sdcla00081' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TMSIS Common Solutions', [Purpose] = N'App Server', [Supportgroup] = N'MMIS Support Data Warehouse TMSIS' WHERE Servername = N'ah57sdcla00082' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TMSIS Common Solutions', [Purpose] = N'App Server', [Supportgroup] = N'MMIS Support Data Warehouse TMSIS' WHERE Servername = N'ah57sdcla00083' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet Enterprise', [Purpose] = N'Filenet CPE Web 1 - OS patching can restart without app admin assistance.', [Supportgroup] = N'MMIS Support Backend Apps Filenet' WHERE Servername = N'ah57sdcla00095' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet', [Purpose] = N'Filenet CPE Web 2 - OS patching can restart without app admin assistance.', [Supportgroup] = N'MMIS Support Backend Apps Filenet' WHERE Servername = N'ah57sdcla00096' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TMSIS Common Solutions', [Purpose] = N'App Server', [Supportgroup] = N'MMIS Support Data Warehouse TMSIS' WHERE Servername = N'ah57sdcla00102' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Code Management', [Purpose] = N'RH Build Box', [Supportgroup] = N'MMIS Support Backend Apps DevOps' WHERE Servername = N'ah57sdcla00112' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Code Management', [Purpose] = N'Scanning', [Supportgroup] = N'MMIS Support Backend Apps DevOps' WHERE Servername = N'ah57sdcla00113' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Code Management', [Purpose] = N'RH Build Box', [Supportgroup] = N'MMIS Support Backend Apps DevOps' WHERE Servername = N'ah57sdcla00114' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Code Management', [Purpose] = N'Scanning', [Supportgroup] = N'MMIS Support Backend Apps DevOps' WHERE Servername = N'ah57sdcla00115' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA DEV (RHEL8)', [Purpose] = N'IBM ACE', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00127' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA DEV (RHEL8)', [Purpose] = N'IBM WAS', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00128' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA DEV (RHEL8)', [Purpose] = N'IBM BAW', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00129' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA DEV (RHEL8)', [Purpose] = N'Progress Corticon', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00130' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA DEV (RHEL8)', [Purpose] = N'IBMWSRR', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00131' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Sandbox (RHEL8)', [Purpose] = N'SOA Sandbox', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00132' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA TEST (RHEL 8)', [Purpose] = N'IBM ACE', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00133' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA TEST (RHEL 8)', [Purpose] = N'IBM WAS', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00134' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA TEST (RHEL 8)', [Purpose] = N'IBM BAW', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00135' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA TEST (RHEL 8)', [Purpose] = N'Progress Corticon', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00136' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA TEST (RHEL 8)', [Purpose] = N'IBMWSRR', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00137' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA UAT (RHEL8)', [Purpose] = N'IBM ACE', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00138' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA UAT (RHEL8)', [Purpose] = N'IBM WAS', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00139' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA UAT (RHEL8)', [Purpose] = N'IBM BAW', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00140' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA UAT (RHEL8)', [Purpose] = N'Progress Corticon', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00141' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA UAT (RHEL8)', [Purpose] = N'IBMWSRR', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00142' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Managed File Transfer', [Purpose] = N'MFT - Hosts the SFG application for managed file transfer', [Supportgroup] = N'MMIS Support EDI SFG/Sterling' WHERE Servername = N'ah57sdcla00143' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Managed File Transfer', [Purpose] = N'MFT - Hosts the SFG application for managed file transfer', [Supportgroup] = N'MMIS Support EDI SFG/Sterling' WHERE Servername = N'ah57sdcla00144' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Managed File Transfer', [Purpose] = N'MFT - Hosts the SFG application for managed file transfer', [Supportgroup] = N'MMIS Support EDI SFG/Sterling' WHERE Servername = N'ah57sdcla00145' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Managed File Transfer', [Purpose] = N'MFT - Hosts the SFG application for managed file transfer', [Supportgroup] = N'MMIS Support EDI SFG/Sterling' WHERE Servername = N'ah57sdcla00146' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Autosys', [Purpose] = N'Autosys App', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcla00147' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Autosys', [Purpose] = N'Autosys App', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcla00148' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Autosys', [Purpose] = N'Autosys App', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcla00149' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Autosys', [Purpose] = N'Autosys App', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcla00150' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Autosys', [Purpose] = N'Autosys App', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcla00151' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Autosys', [Purpose] = N'Autosys App', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcla00152' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage (RHEL8)', [Purpose] = N'IBM ACE', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00153' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage (RHEL8)', [Purpose] = N'IBM ACE', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00154' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage (RHEL8)', [Purpose] = N'IBM WAS', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00155' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage (RHEL8)', [Purpose] = N'IBM WAS', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00156' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage (RHEL8)', [Purpose] = N'IBM BAW', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00157' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage (RHEL8)', [Purpose] = N'IBM BAW', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00158' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage (RHEL8)', [Purpose] = N'Progress Corticon', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00159' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage (RHEL8)', [Purpose] = N'Progress Corticon', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00160' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage (RHEL8)', [Purpose] = N'IBM WSRR', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00161' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage (RHEL8)', [Purpose] = N'IBM WSRR', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00162' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TMSIS Common Solutions', [Purpose] = N'App Server', [Supportgroup] = N'MMIS Support Data Warehouse TMSIS' WHERE Servername = N'ah57sdcla00163' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TMSIS Common Solutions', [Purpose] = N'App Server', [Supportgroup] = N'MMIS Support Data Warehouse TMSIS' WHERE Servername = N'ah57sdcla00164' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TMSIS Common Solutions', [Purpose] = N'App Server', [Supportgroup] = N'MMIS Support Data Warehouse TMSIS' WHERE Servername = N'ah57sdcla00165' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TMSIS Common Solutions', [Purpose] = N'App Server', [Supportgroup] = N'MMIS Support Data Warehouse TMSIS' WHERE Servername = N'ah57sdcla00166' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TMSIS Common Solutions', [Purpose] = N'App Server', [Supportgroup] = N'MMIS Support Data Warehouse TMSIS' WHERE Servername = N'ah57sdcla00167' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'IBM ACE', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00168' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'IBM ACE', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00169' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'IBM WAS', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00170' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'IBM WAS', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00171' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'IBM BAW', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00172' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'IBM BAW', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00173' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'Progress Corticon', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00174' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'Progress Corticon', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00175' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'IBM WSRR', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00176' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'IBM WSRR', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00177' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'IBM License Metric Tool (IBM LMT)', [Purpose] = N'IBM Licensing Server', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla00178' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (IIB)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla01050' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (IIB)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla01051' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (WAS)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla01052' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (WAS)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla01053' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (BPM)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla01054' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (BPM)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla01055' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (WAS)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla01056' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (WAS)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla01057' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (WSRR)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla01058' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Application Server (WSRR)', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcla01059' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Managed File Transfer', [Purpose] = N'MFT - Hosts the SFG application for managed file transfer', [Supportgroup] = N'MMIS Support EDI SFG/Sterling' WHERE Servername = N'ah57sdcla03023' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Managed File Transfer', [Purpose] = N'MFT - Hosts the SFG application for managed file transfer', [Supportgroup] = N'MMIS Support EDI SFG/Sterling' WHERE Servername = N'ah57sdcla03024' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Managed File Transfer', [Purpose] = N'MFT - Hosts the SFG application for managed file transfer', [Supportgroup] = N'MMIS Support EDI SFG/Sterling' WHERE Servername = N'ah57sdcla03025' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Managed File Transfer', [Purpose] = N'MFT - Hosts the SFG application for managed file transfer', [Supportgroup] = N'MMIS Support EDI SFG/Sterling' WHERE Servername = N'ah57sdcla03026' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS Dynatrace', [Purpose] = N'MMIS Dynatrace Node', [Supportgroup] = N'MMIS DynaTrace SW ADMIN' WHERE Servername = N'ah57sdcla04201' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS Dynatrace', [Purpose] = N'MMIS Dynatrace Node', [Supportgroup] = N'MMIS DynaTrace SW ADMIN' WHERE Servername = N'ah57sdcla04202' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS Dynatrace', [Purpose] = N'MMIS Dynatrace Node', [Supportgroup] = N'MMIS DynaTrace SW ADMIN' WHERE Servername = N'ah57sdcla04203' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS Dynatrace', [Purpose] = N'MMIS DynaTrace Cluster Active gate', [Supportgroup] = N'MMIS DynaTrace SW ADMIN' WHERE Servername = N'ah57sdcla04204' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS Dynatrace', [Purpose] = N'MMIS DynaTrace Cluster Active gate', [Supportgroup] = N'MMIS DynaTrace SW ADMIN' WHERE Servername = N'ah57sdcla04205' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS Dynatrace', [Purpose] = N'MMIS DynaTrace Environment Active Gate', [Supportgroup] = N'MMIS DynaTrace SW ADMIN' WHERE Servername = N'ah57sdcla04206' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS Dynatrace', [Purpose] = N'MMIS DynaTrace Environment Active Gate', [Supportgroup] = N'MMIS DynaTrace SW ADMIN' WHERE Servername = N'ah57sdcla04207' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [Purpose] = N'NetBackup' WHERE Servername = N'ah57sdcla90013' AND ([Purpose] IS NULL);
UPDATE #TempSnowServer SET [Purpose] = N'NetBackup' WHERE Servername = N'ah57sdcla90014' AND ([Purpose] IS NULL);
UPDATE #TempSnowServer SET [Purpose] = N'NetBackup' WHERE Servername = N'ah57sdcla90015' AND ([Purpose] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Edifecs', [Purpose] = N'Oracle DB', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdclb03003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Edifecs', [Purpose] = N'Oracle DB', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdclb03004' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ATP', [Purpose] = N'Oracle RAC Node - Uses IPTables', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdclc00001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ATP', [Purpose] = N'Oracle RAC Node - Uses IPTables', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdclc00002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Shared RAC', [Purpose] = N'Oracle RAC Node', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdclc00007' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Shared RAC', [Purpose] = N'Oracle RAC Node', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdclc00008' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle RAC', [Purpose] = N'Oracle RAC Node 1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdclc03003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle RAC', [Purpose] = N'Oracle RAC Node 2', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdclc03004' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle RAC - Sandbox', [Purpose] = N'Oracle RAC Sandbox Node 1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdclc03005' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle RAC - Sandbox', [Purpose] = N'Oracle RAC Sandbox Node 2', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdclc03006' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle RAC', [Purpose] = N'Oracle RAC Node 1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdclc03007' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle RAC', [Purpose] = N'Oracle RAC Node 2', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdclc03008' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle RAC', [Purpose] = N'Oracle RAC Node 1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdclc03009' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle RAC', [Purpose] = N'Oracle RAC Node 2', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdclc03010' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle RAC', [Purpose] = N'Oracle RAC Node 1', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdclc04003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle RAC', [Purpose] = N'Oracle RAC Node 2', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdclc04004' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA', [Purpose] = N'Datapower', [Supportgroup] = N'ENTERPRISE SOA SW ADMIN' WHERE Servername = N'ah57sdcli00002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA', [Purpose] = N'Datapower', [Supportgroup] = N'ENTERPRISE SOA SW ADMIN' WHERE Servername = N'ah57sdcli00003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Syslog', [Purpose] = N'Syslog - OS patching can restart without app admin assistance', [Supportgroup] = N'ATP SW ADMIN' WHERE Servername = N'ah57sdclm00001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Code Management', [Purpose] = N'Code Management Reposity  Server - OS patching can restart without app admin assistance.', [Supportgroup] = N'MMIS Support Backend Apps DevOps' WHERE Servername = N'ah57sdclw00060' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Code Management', [Purpose] = N'Code Management Reposity  Server - OS patching can restart without app admin assistance.', [Supportgroup] = N'MMIS Support Backend Apps DevOps' WHERE Servername = N'ah57sdclw00061' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Web Server', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdclw00096' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'SOA Web Server', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdclw00097' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Web/App/Rpt - can be restarted without app admin assistance', [Supportgroup] = N'SIEBEL SW ADMIN' WHERE Servername = N'ah57sdclw00103' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Code Management', [Purpose] = N'Repository', [Supportgroup] = N'MMIS Support Backend Apps DevOps' WHERE Servername = N'ah57sdclw00107' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Code Management', [Purpose] = N'Repository', [Supportgroup] = N'MMIS Support Backend Apps DevOps' WHERE Servername = N'ah57sdclw00108' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage (RHEL8)', [Purpose] = N'TMED Web Server', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdclw00109' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Stage (RHEL8)', [Purpose] = N'TMED Web Server', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdclw00110' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'TMED Web Server', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdclw00111' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR (RHEL8)', [Purpose] = N'TMED Web Server', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdclw00112' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'AGM', [Purpose] = N'AGM Web Node 1 - OS patching can restart without app admin assistance', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdclw20002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'OEM', [Purpose] = N'Oracle Enterprise Manager', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdclx00002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle Enterprise Manager (OEM)', [Purpose] = N'Oracle Enterprise Manager', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdclx00004' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Oracle Enterprise Manager (OEM)', [Purpose] = N'OEM Server', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdclx00005' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'PAM', [Purpose] = N'PAM Appliance', [Supportgroup] = N'PAM SW Admin' WHERE Servername = N'ah57sdcpa00003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'PAM', [Purpose] = N'PAM Appliance', [Supportgroup] = N'PAM SW Admin' WHERE Servername = N'ah57sdcpa00004' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'PAM', [Purpose] = N'PAM Debian appliance', [Supportgroup] = N'UNIX OS ADMIN' WHERE Servername = N'ah57sdcpa00005' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA', [Purpose] = N'API Connect 2018.04.1', [Supportgroup] = N'ENTERPRISE SOA SW ADMIN' WHERE Servername = N'ah57sdcpa00008' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA', [Purpose] = N'API Connect 2018.04.1', [Supportgroup] = N'ENTERPRISE SOA SW ADMIN' WHERE Servername = N'ah57sdcpa00009' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA', [Purpose] = N'API Connect 2018.04.1', [Supportgroup] = N'ENTERPRISE SOA SW ADMIN' WHERE Servername = N'ah57sdcpa00010' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'Datapower', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcpi00005-c' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SOA Production-DR', [Purpose] = N'DataPower Physical Gateway S2', [Supportgroup] = N'MMIS Support SOA - SOA' WHERE Servername = N'ah57sdcpi00006-c' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SAS-BI', [Purpose] = N'Application Server', [Supportgroup] = N'SAS SW Admin' WHERE Servername = N'ah57sdcua00001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SAS-BI', [Purpose] = N'SAS Application Server', [Supportgroup] = N'SAS SW Admin' WHERE Servername = N'ah57sdcua00002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Interchange', [Purpose] = N'Oracle database - Filenet/Siebel/Exstream - idbp8p sblprod tnedrp1 tnedrp2', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub00016' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Interchange', [Purpose] = N'Oracle Financial', [Supportgroup] = N'UNIX OS ADMIN' WHERE Servername = N'ah57sdcub01018' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Purpose] = N'Oracle db - MMIS/DW/MAR/Fraud & Abuse - ...', [Supportgroup] = N'UNIX OS ADMIN' WHERE Servername = N'ah57sdcub01020' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Interchange', [Purpose] = N'Oracle database', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub01021' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TPAES', [Purpose] = N'TPAES 2 Development Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub01027' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM', [Purpose] = N'Oracle Database Domain on M8000. Requested to be operational by Chad Williams at 9:14AM CST on 9/15/2023', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub02001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM', [Purpose] = N'Oracle Database domain on M8000', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcub02001- dup' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcub04003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04004' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'MMIS Support Backend Apps DevOps' WHERE Servername = N'ah57sdcub04005' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdcub04006' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database - MMIS', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04007' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - Siebel/CMOD/FileNet - Production - SBLPRD20, ARS105P, IBMCNVP, IBMFNVP, IDBP55P', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdcub04008' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'ForgeRock LDAP integration for interchange', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcub04009' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04010' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Code Management', [Purpose] = N'Code Management Build Server - OS patching can restart without app admin assistance.', [Supportgroup] = N'MMIS Support Backend Apps DevOps' WHERE Servername = N'ah57sdcub04011' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04012' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04013' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database - Controlled Substance - tncsmp1', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04014' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database - Controlled Substance Load  - tncsma2', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04015' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - CSM - Production - CSMSP', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdcub04016' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - CSM - Production - CSMP', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdcub04017' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - Siebel - Production - SBLPRD17', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdcub04018' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04019' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone Duplicate of Ah57sdcub04020 10.11.36.120', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcub04020' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcub04021' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'ForgeRock LDAP', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcub04022' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database - MMIS DW', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04023' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T-7 Database Host', [Purpose] = N'Oracle Database Zone - ALM/DataStage - Production - TNALMP2, TNDSTGP3', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdcub04024' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04025' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04026' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04027' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04028' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04029' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcub04030' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04031' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04032' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04033' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04034' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04035' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04036' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04037' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04038' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04039' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database', [Purpose] = N'Oracle Database Zone', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcub04040' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SFTP', [Purpose] = N'LDOM/Zone', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcub04041' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SFTP', [Purpose] = N'LDOM/Zone', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcub04043' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SFTP', [Purpose] = N'LDOM/Zone', [Supportgroup] = N'Legacy SFTP SW ADMIN' WHERE Servername = N'ah57sdcub04045' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Code Management', [Purpose] = N'Solaris Build Box - invalid host. Retiring 9/22/21', [Supportgroup] = N'CODE MANAGEMENT SW ADMIN' WHERE Servername = N'ah57sdcub04046' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SFTP', [Purpose] = N'LDOM / Zone', [Supportgroup] = N'Legacy SFTP SW ADMIN' WHERE Servername = N'ah57sdcub04048' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SFTP', [Purpose] = N'LDOM/Zone', [Supportgroup] = N'Legacy SFTP SW ADMIN' WHERE Servername = N'ah57sdcub04049' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T-8 Database Host', [Purpose] = N'Interchange prod databases - MMIS - TNHISP, TNMISP, TNDOCP', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdcub04056' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T-8 Database Host', [Purpose] = N'Interchange Data Warehouse - MMIS - TNMARP, TNSURP', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdcub04057' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T8 Database Host', [Purpose] = N'Oracle Database - SAS - Production - SASP', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdcub04060' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T8 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - DEHIS*, DEMIS*', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdcub04061' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T-7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - TNHISA, TNMISA', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdcub04062' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Zone - MMIS - DEHISA, DEMISA', [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdcub04063' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdcub04064' AND ([Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [Supportgroup] = N'MMIS Support INF DBA' WHERE Servername = N'ah57sdcub04066' AND ([Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Host', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcuh04001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T7 Database Host', [Purpose] = N'Oracle Database Host', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcuh04002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SFTP', [Purpose] = N'LDOM/Zone Host', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcuh04003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SFTP', [Purpose] = N'LDOM/Zone Host', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcuh04004' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'T-8 Database Host', [Purpose] = N'T8 Oracle Host', [Supportgroup] = N'MMIS Support INF UNIX' WHERE Servername = N'ah57sdcuh04005' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Purpose] = N'Oracle db - MMIS - tnhisp1 tnjobp1 tnjobp2 tnjobp3 tnmisp1', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcux01013' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Purpose] = N'Oracle db - MAR/Fraud & Abuse/BO/Doco - rcatprod tndocp1 tnmarp1 tnsump1 tnsurp1', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcux01015' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Purpose] = N'Oracle db - MMIS/DW/MAR/Fraud & Abuse - t9hisa1  t9mara1 t9misa1 t9sura1', [Supportgroup] = N'Unix Admin' WHERE Servername = N'ah57sdcux01017' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Purpose] = N'Oracle database', [Supportgroup] = N'Unix Admin' WHERE Servername = N'ah57sdcux01025' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'OEM', [Purpose] = N'Oracle Database', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcux01027' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [Supportgroup] = N'Unix Admin' WHERE Servername = N'ah57sdcux01028' AND ([Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [Purpose] = N'8:31 AM 4/1/2025 Retired per Chad Williams - Anthony Bronsdon', [Supportgroup] = N'Unix Admin' WHERE Servername = N'ah57sdcux01029' AND ([Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcux01030' AND ([System] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcux01031' AND ([System] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcux01033' AND ([System] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcux01034' AND ([System] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcux01035' AND ([System] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcux01036' AND ([System] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Interchange', [Purpose] = N'Oracle database', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcux01037' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'InterChange', [Purpose] = N'Oracle database', [Supportgroup] = N'ORACLE DB ADMIN DXC' WHERE Servername = N'ah57sdcux01038' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Interchange', [Purpose] = N'Oracle database', [Supportgroup] = N'Unix Admin' WHERE Servername = N'ah57sdcux01041' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Code Management', [Purpose] = N'Code Management - OS patching can restart without app admin assistance.', [Supportgroup] = N'MMIS Support Backend Apps DevOps' WHERE Servername = N'ah57sdcwa00051' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel 2016 App', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00056' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel 2016 App', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00057' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel 2016 App', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00058' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD', [Purpose] = N'Webi Server', [Supportgroup] = N'CMOD SW ADMIN' WHERE Servername = N'ah57sdcwa00066' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD', [Purpose] = N'Webi Server', [Supportgroup] = N'CMOD SW ADMIN' WHERE Servername = N'ah57sdcwa00067' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD', [Purpose] = N'Library/Object Server', [Supportgroup] = N'CMOD SW ADMIN' WHERE Servername = N'ah57sdcwa00068' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD', [Purpose] = N'Library/Object Server', [Supportgroup] = N'CMOD SW ADMIN' WHERE Servername = N'ah57sdcwa00069' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD', [Purpose] = N'Webi Server', [Supportgroup] = N'CMOD SW ADMIN' WHERE Servername = N'ah57sdcwa00070' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD', [Purpose] = N'Library/Object Server', [Supportgroup] = N'CMOD SW ADMIN' WHERE Servername = N'ah57sdcwa00071' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD', [Purpose] = N'Webi Server', [Supportgroup] = N'CMOD SW ADMIN' WHERE Servername = N'ah57sdcwa00072' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD', [Purpose] = N'Webi Server', [Supportgroup] = N'CMOD SW ADMIN' WHERE Servername = N'ah57sdcwa00073' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet Enterprise', [Purpose] = N'Filenet CPE Web 1', [Supportgroup] = N'FILENET SW ADMIN DXC' WHERE Servername = N'ah57sdcwa00074' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet Enterprise', [Purpose] = N'Filenet CPE Web 2', [Supportgroup] = N'FILENET SW ADMIN DXC' WHERE Servername = N'ah57sdcwa00075' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet Enterprise', [Purpose] = N'Filenet ICN Web 1', [Supportgroup] = N'FILENET SW ADMIN DXC' WHERE Servername = N'ah57sdcwa00076' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet Enterprise', [Purpose] = N'Filenet ICN Web 2', [Supportgroup] = N'FILENET SW ADMIN DXC' WHERE Servername = N'ah57sdcwa00077' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet Enterprise', [Purpose] = N'Filenet ICC', [Supportgroup] = N'FILENET SW ADMIN DXC' WHERE Servername = N'ah57sdcwa00078' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet Enterprise', [Purpose] = N'Filenet IDL', [Supportgroup] = N'FILENET SW ADMIN DXC' WHERE Servername = N'ah57sdcwa00079' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet Enterprise', [Purpose] = N'Filenet CSS', [Supportgroup] = N'FILENET SW ADMIN DXC' WHERE Servername = N'ah57sdcwa00080' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet Enterprise', [Purpose] = N'Filenet CPE/ICN Web 1', [Supportgroup] = N'MMIS Support Backend Apps Filenet' WHERE Servername = N'ah57sdcwa00081' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet Enterprise', [Purpose] = N'Filenet CPE/ICN Web 2', [Supportgroup] = N'FILENET SW ADMIN DXC' WHERE Servername = N'ah57sdcwa00082' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet Enterprise', [Purpose] = N'Filenet ICC', [Supportgroup] = N'FILENET SW ADMIN DXC' WHERE Servername = N'ah57sdcwa00083' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet Enterprise', [Purpose] = N'Filenet IDL', [Supportgroup] = N'FILENET SW ADMIN DXC' WHERE Servername = N'ah57sdcwa00084' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet Enterprise', [Purpose] = N'Filenet CSS', [Supportgroup] = N'FILENET SW ADMIN DXC' WHERE Servername = N'ah57sdcwa00085' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet Enterprise', [Purpose] = N'Filenet CPE/ICN Web 1', [Supportgroup] = N'FILENET SW ADMIN DXC' WHERE Servername = N'ah57sdcwa00086' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet Enterprise', [Purpose] = N'Filenet CPE/ICN Web 2', [Supportgroup] = N'FILENET SW ADMIN DXC' WHERE Servername = N'ah57sdcwa00087' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet Enterprise', [Purpose] = N'Filenet ICC', [Supportgroup] = N'FILENET SW ADMIN DXC' WHERE Servername = N'ah57sdcwa00088' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Filenet Enterprise', [Purpose] = N'Filenet IDL', [Supportgroup] = N'FILENET SW ADMIN DXC' WHERE Servername = N'ah57sdcwa00089' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Filenet Enterprise', [Purpose] = N'Filenet CSS', [Supportgroup] = N'FILENET SW ADMIN DXC' WHERE Servername = N'ah57sdcwa00090' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'PAM', [Purpose] = N'CA Authentication', [Supportgroup] = N'PAM SW Admin' WHERE Servername = N'ah57sdcwa00095' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Code Management', [Purpose] = N'Win Build Box - OS patching can restart without app admin assistance.', [Supportgroup] = N'MMIS Support Backend Apps DevOps' WHERE Servername = N'ah57sdcwa00096' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel 2016 Prod app', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00097' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Code Management', [Purpose] = N'Code Management Build Server - OS patching can restart without app admin assistance.', [Supportgroup] = N'MMIS Support Backend Apps DevOps' WHERE Servername = N'ah57sdcwa00100' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Code Management', [Purpose] = N'Code Management Build Server - OS patching can restart without app admin assistance.', [Supportgroup] = N'MMIS Support Backend Apps DevOps' WHERE Servername = N'ah57sdcwa00101' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'JIL Manager (Job Information Language)', [Purpose] = N'Job Information Language Application Server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcwa00102' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'JIL Manager (Job Information Language)', [Purpose] = N'Job Information Language Application Server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcwa00103' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Web/App', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00104' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Web/App', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00105' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Application', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00106' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Application', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00107' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Application', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00108' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Reports/App', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00109' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Web/App', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00110' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Web/App', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00111' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Application', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00112' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Application', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00113' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Application', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00114' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Reports/App', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00115' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Web/App', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00116' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Web/App', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00117' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Application', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00118' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Application', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00119' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Application', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00120' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Reports/App', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00121' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Web/App/Reports', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00122' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel Web/App/Report', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcwa00123' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Site Scope', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00124' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise (old)', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00125' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM â€“ Application Lifecycle Management', [Purpose] = N'ALM Synchronizer', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00126' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM â€“ Application Lifecycle Management', [Purpose] = N'MFConnect', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00127' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Diag', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00128' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM â€“ Application Lifecycle Management', [Purpose] = N'UFT Licensing Server (Autopass)', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00129' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Host 1', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00130' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Host 2', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00131' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Host 3', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00132' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Host 4', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00133' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Host 5', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00134' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Host 6', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00135' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Host 7', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00136' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Host 8', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00137' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Host 9', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00138' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Host 10', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00139' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Host 11', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00140' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Host 12', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00141' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Host 13', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00142' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Host 14', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00143' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Host 15', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00144' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Site Scope', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00145' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise (old)', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00146' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM â€“ Application Lifecycle Management', [Purpose] = N'ALM Synchronizer', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00147' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM â€“ Application Lifecycle Management', [Purpose] = N'MFConnect', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00148' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Diag', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00149' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TPrime', [Purpose] = N'Search/Application', [Supportgroup] = N'TPRIME SW ADMIN' WHERE Servername = N'ah57sdcwa00150' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM â€“ Application Lifecycle Management', [Purpose] = N'UFT Licensing Server (Autopass)', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00152' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD â€“ Content Management On Demand', [Purpose] = N'IBM Content Navigator', [Supportgroup] = N'MMIS Support Backend Apps CMOD' WHERE Servername = N'ah57sdcwa00154' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD â€“ Content Management On Demand', [Purpose] = N'IBM Content Navigator', [Supportgroup] = N'MMIS Support Backend Apps CMOD' WHERE Servername = N'ah57sdcwa00155' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD â€“ Content Management On Demand', [Purpose] = N'Library / Object Server', [Supportgroup] = N'MMIS Support Backend Apps CMOD' WHERE Servername = N'ah57sdcwa00156' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD â€“ Content Management On Demand', [Purpose] = N'Library / Object Server', [Supportgroup] = N'MMIS Support Backend Apps CMOD' WHERE Servername = N'ah57sdcwa00157' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD â€“ Content Management On Demand', [Purpose] = N'IBM Content Navigator', [Supportgroup] = N'MMIS Support Backend Apps CMOD' WHERE Servername = N'ah57sdcwa00158' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD â€“ Content Management On Demand', [Purpose] = N'Library / Object Server', [Supportgroup] = N'MMIS Support Backend Apps CMOD' WHERE Servername = N'ah57sdcwa00159' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD â€“ Content Management On Demand', [Purpose] = N'IBM Content Navigator', [Supportgroup] = N'MMIS Support Backend Apps CMOD' WHERE Servername = N'ah57sdcwa00160' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD â€“ Content Management On Demand', [Purpose] = N'IBM Content Navigator', [Supportgroup] = N'MMIS Support Backend Apps CMOD' WHERE Servername = N'ah57sdcwa00161' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD â€“ Content Management On Demand', [Purpose] = N'Library / Object Server', [Supportgroup] = N'MMIS Support Backend Apps CMOD' WHERE Servername = N'ah57sdcwa00162' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD â€“ Content Management On Demand', [Purpose] = N'IBM Content Navigator', [Supportgroup] = N'MMIS Support Backend Apps CMOD' WHERE Servername = N'ah57sdcwa00163' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer App 1', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcwa00164' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer App 2', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcwa00165' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer App 3', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcwa00166' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer App 1', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcwa00167' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer App 2', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcwa00168' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer App 3', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcwa00169' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Host 1', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00170' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise Host 2', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcwa00171' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet', [Purpose] = N'Filenet ICC', [Supportgroup] = N'MMIS Support Backend Apps Filenet' WHERE Servername = N'ah57sdcwa00172' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet', [Purpose] = N'Filenet IDL', [Supportgroup] = N'MMIS Support Backend Apps Filenet' WHERE Servername = N'ah57sdcwa00173' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet', [Purpose] = N'Filenet CSS', [Supportgroup] = N'MMIS Support Backend Apps Filenet' WHERE Servername = N'ah57sdcwa00174' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet', [Purpose] = N'Filenet ICC', [Supportgroup] = N'MMIS Support Backend Apps Filenet' WHERE Servername = N'ah57sdcwa00175' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet', [Purpose] = N'Filenet IDL/CSS', [Supportgroup] = N'MMIS Support Backend Apps Filenet' WHERE Servername = N'ah57sdcwa00176' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet', [Purpose] = N'Filenet ICC', [Supportgroup] = N'MMIS Support Backend Apps Filenet' WHERE Servername = N'ah57sdcwa00177' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet', [Purpose] = N'Filenet IDL/CSS', [Supportgroup] = N'MMIS Support Backend Apps Filenet' WHERE Servername = N'ah57sdcwa00178' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Exstream', [Purpose] = N'App Server', [Supportgroup] = N'MMIS Support MMIS Front End Letters' WHERE Servername = N'ah57sdcwa00179' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Test Data Manager', [Purpose] = N'Data Masking 1', [Supportgroup] = N'TDM SW ADMIN' WHERE Servername = N'ah57sdcwa00180' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Test Data Manager', [Purpose] = N'Data Masking 2', [Supportgroup] = N'TDM SW ADMIN' WHERE Servername = N'ah57sdcwa00181' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Exstream', [Purpose] = N'Command Center', [Supportgroup] = N'MMIS Support MMIS Front End Letters' WHERE Servername = N'ah57sdcwa00182' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Exstream', [Purpose] = N'Licensing Server', [Supportgroup] = N'MMIS Support MMIS Front End Letters' WHERE Servername = N'ah57sdcwa00183' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Exstream', [Purpose] = N'Command Center', [Supportgroup] = N'MMIS Support MMIS Front End Letters' WHERE Servername = N'ah57sdcwa00184' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Exstream', [Purpose] = N'Command Center', [Supportgroup] = N'MMIS Support MMIS Front End Letters' WHERE Servername = N'ah57sdcwa00185' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Exstream', [Purpose] = N'Licensing Server', [Supportgroup] = N'MMIS Support MMIS Front End Letters' WHERE Servername = N'ah57sdcwa00186' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'RRI - Resource Recognition Incorporated', [Purpose] = N'Qmanager App', [Supportgroup] = N'MMIS Support Backend Apps Formworks' WHERE Servername = N'ah57sdcwa00187' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'RRI - Resource Recognition Incorporated', [Purpose] = N'BG Worker 1 App', [Supportgroup] = N'MMIS Support Backend Apps Formworks' WHERE Servername = N'ah57sdcwa00188' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'RRI - Resource Recognition Incorporated', [Purpose] = N'BG Worker 2 App', [Supportgroup] = N'MMIS Support Backend Apps Formworks' WHERE Servername = N'ah57sdcwa00189' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'RRI - Resource Recognition Incorporated', [Purpose] = N'WebKey / Bridges 1', [Supportgroup] = N'MMIS Support Backend Apps Formworks' WHERE Servername = N'ah57sdcwa00190' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'RRI - Resource Recognition Incorporated', [Purpose] = N'Output QA Store', [Supportgroup] = N'MMIS Support Backend Apps Formworks' WHERE Servername = N'ah57sdcwa00191' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'RRI - Resource Recognition Incorporated', [Purpose] = N'Lookup', [Supportgroup] = N'MMIS Support Backend Apps Formworks' WHERE Servername = N'ah57sdcwa00192' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Decision Analytics', [Purpose] = N'Qlik Server / Database replication to cloud', [Supportgroup] = N'MMIS Support Decision Analytics Dashboard' WHERE Servername = N'ah57sdcwa00197' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Decision Analytics', [Purpose] = N'Qlik Server / Database replication to cloud', [Supportgroup] = N'MMIS Support Decision Analytics Dashboard' WHERE Servername = N'ah57sdcwa00198' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer App 1', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcwa00202' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer App 2', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcwa00203' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer App1', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcwa00204' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer App2', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcwa00205' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Exstream', [Purpose] = N'Print Server', [Supportgroup] = N'EXSTREAM SW ADMIN' WHERE Servername = N'ah57sdcwa01137' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Performance Center', [Purpose] = N'ALM PC Host', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20008' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Performance Center', [Purpose] = N'ALM PC Host', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20009' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Performance Center', [Purpose] = N'ALM PC Host', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20010' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Performance Center', [Purpose] = N'ALM PC Host', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20011' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Performance Center', [Purpose] = N'ALM PC Host', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20012' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Performance Center', [Purpose] = N'ALM PC Host', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20013' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Performance Center', [Purpose] = N'ALM PC Host', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20014' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Performance Center', [Purpose] = N'ALM PC Host', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20015' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Performance Center', [Purpose] = N'ALM PC Host', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20016' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Performance Center', [Purpose] = N'ALM PC Host', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20017' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Performance Center', [Purpose] = N'ALM PC Host', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20018' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Performance Center', [Purpose] = N'ALM PC Host', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20019' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Performance Center', [Purpose] = N'ALM PC Host', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20020' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Performance Center', [Purpose] = N'ALM PC Host', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20021' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Performance Center', [Purpose] = N'ALM PC Host', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20022' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM', [Purpose] = N'SiteScope', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20023' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM', [Purpose] = N'ALM Sync', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20024' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'AGM', [Purpose] = N'AGM Sync', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20025' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM', [Purpose] = N'ALM PC Diag', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20026' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM', [Purpose] = N'ALM UFT Licensing Server', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcwa20027' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Jira for MMIS', [Purpose] = N'Jira Web Server', [Supportgroup] = N'AGM SW ADMIN' WHERE Servername = N'ah57sdcwa20028' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TPrime SharePoint SQL', [Purpose] = N'SharePoint Database Server', [Supportgroup] = N'MMIS Support Backend Apps SharePoint' WHERE Servername = N'ah57sdcwb00002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Sharepoint Qlik', [Purpose] = N'Qlik Prod  and Sharepoint Migration tool', [Supportgroup] = N'MMIS Change Backend Apps Sharepoint' WHERE Servername = N'ah57sdcwb00013' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TEAMS', [Purpose] = N'MS SQL DB Server', [Supportgroup] = N'TEAMS SW ADMIN' WHERE Servername = N'ah57sdcwb00019' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TEAMS', [Purpose] = N'MS SQL DB Server', [Supportgroup] = N'TEAMS SW ADMIN' WHERE Servername = N'ah57sdcwb00020' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TPrime', [Purpose] = N'SQL DB Server', [Supportgroup] = N'TPRIME SW ADMIN' WHERE Servername = N'ah57sdcwb00021' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TPrime', [Purpose] = N'T-Prime SharePoint SQL DB Server', [Supportgroup] = N'MMIS Change Backend Apps Sharepoint' WHERE Servername = N'ah57sdcwb00027' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TPrime', [Purpose] = N'MS SQL DB Server', [Supportgroup] = N'TPRIME SW ADMIN' WHERE Servername = N'ah57sdcwb00040' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TPrime', [Purpose] = N'MS SQL DB Server', [Supportgroup] = N'TPRIME SW ADMIN' WHERE Servername = N'ah57sdcwb00041' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Decision Analytics', [Purpose] = N'UAT / STAGING SQL Enterprise & SSIS', [Supportgroup] = N'WINDOWS OS ADMIN STS' WHERE Servername = N'ah57sdcwb00042' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Decision Analytics', [Purpose] = N'DEV SQL Enterprise & SSIS', [Supportgroup] = N'WINDOWS OS ADMIN STS' WHERE Servername = N'ah57sdcwb00043' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Decision Analytics', [Purpose] = N'Prod Cluster 1 Power BI, SSDS & SSAS', [Supportgroup] = N'WINDOWS OS ADMIN STS' WHERE Servername = N'ah57sdcwc00001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Decision Analytics', [Purpose] = N'Prod Power BI, SSDS & SSAS', [Supportgroup] = N'WINDOWS OS ADMIN STS' WHERE Servername = N'ah57sdcwc00002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Decision Analytics', [Purpose] = N'Prod Cluster 2 SQL Ent & SSIS', [Supportgroup] = N'WINDOWS OS ADMIN STS' WHERE Servername = N'ah57sdcwc00003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Decision Analytics', [Purpose] = N'Prod Cluster 2 SQL Ent & SSIS', [Supportgroup] = N'WINDOWS OS ADMIN STS' WHERE Servername = N'ah57sdcwc00004' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileShares', [Purpose] = N'Windows File Server', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'ah57sdcwf00001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Decision Analytics', [Purpose] = N'Terminal Services / Prod Support / UAT Testing', [Supportgroup] = N'MMIS Support Decision Analytics Dashboard' WHERE Servername = N'ah57sdcwt00001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Decision Analytics', [Purpose] = N'Terminal Services / DEV', [Supportgroup] = N'MMIS Support Decision Analytics Dashboard' WHERE Servername = N'ah57sdcwt00002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TDM', [Purpose] = N'CA Test Data Management', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcww00029' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Datastage', [Purpose] = N'Web/App Server', [Supportgroup] = N'DATASTAGE SW ADMIN' WHERE Servername = N'ah57sdcww00030' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel 2016 Web', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcww00035' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Siebel 2016 Web', [Supportgroup] = N'MMIS Support Backend Apps Siebel' WHERE Servername = N'ah57sdcww00036' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SharePoint Qlik', [Purpose] = N'Qlik Dev & SQL Server and Sharepoint Migration tool', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'ah57sdcww00037' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SharePoint', [Purpose] = N'SharePoint Web', [Supportgroup] = N'MMIS Change Backend Apps Sharepoint' WHERE Servername = N'ah57sdcww00038' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'Interchange Web', [Supportgroup] = N'APPEON SW ADMIN' WHERE Servername = N'ah57sdcww00040' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'Interchange Web', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww00042' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Datastage', [Purpose] = N'Web/App Server', [Supportgroup] = N'DATASTAGE SW ADMIN' WHERE Servername = N'ah57sdcww00044' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Datastage', [Purpose] = N'Web/App Server', [Supportgroup] = N'DATASTAGE SW ADMIN' WHERE Servername = N'ah57sdcww00045' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Datastage', [Purpose] = N'Web/App Server', [Supportgroup] = N'DATASTAGE SW ADMIN' WHERE Servername = N'ah57sdcww00046' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Doctor Lookup', [Purpose] = N'Web Server', [Supportgroup] = N'DRLOOKUP SW ADMIN' WHERE Servername = N'ah57sdcww00053' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Doctor Lookup', [Purpose] = N'Web Server', [Supportgroup] = N'DRLOOKUP SW ADMIN' WHERE Servername = N'ah57sdcww00054' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Doctor Lookup', [Purpose] = N'Web Server', [Supportgroup] = N'DRLOOKUP SW ADMIN' WHERE Servername = N'ah57sdcww00055' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM â€“ Application Lifecycle Management', [Purpose] = N'ALM Web Node 1', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcww00059' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM â€“ Application Lifecycle Management', [Purpose] = N'ALM Web Node 2', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcww00060' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM â€“ Application Lifecycle Management', [Purpose] = N'ALM Web Node 1', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcww00061' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM â€“ Application Lifecycle Management', [Purpose] = N'ALM Web Node 2', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcww00062' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Datastage', [Purpose] = N'App/Web Server', [Supportgroup] = N'MMIS Support Data Warehouse DataStage' WHERE Servername = N'ah57sdcww00063' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Datastage', [Purpose] = N'App/Web Server', [Supportgroup] = N'MMIS Support Data Warehouse DataStage' WHERE Servername = N'ah57sdcww00064' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Datastage', [Purpose] = N'App/Web Server', [Supportgroup] = N'MMIS Support Data Warehouse DataStage' WHERE Servername = N'ah57sdcww00065' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Datastage', [Purpose] = N'App/Web Server', [Supportgroup] = N'MMIS Support Data Warehouse DataStage' WHERE Servername = N'ah57sdcww00066' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Datastage', [Purpose] = N'App/Web Server', [Supportgroup] = N'MMIS Support Data Warehouse DataStage' WHERE Servername = N'ah57sdcww00067' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Datastage', [Purpose] = N'App/Web Server', [Supportgroup] = N'MMIS Support Data Warehouse DataStage' WHERE Servername = N'ah57sdcww00068' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer Web', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww00069' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer Web', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww00070' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise 1', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcww00071' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise 2', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcww00072' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise 1', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcww00073' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LoadRunner Enterprise', [Purpose] = N'LoadRunner Enterprise 2', [Supportgroup] = N'MMIS Support Backend Apps ALM' WHERE Servername = N'ah57sdcww00074' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet', [Purpose] = N'Filenet CPE Web 1', [Supportgroup] = N'MMIS Support Backend Apps Filenet' WHERE Servername = N'ah57sdcww00075' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet', [Purpose] = N'Filenet CPE Web 2', [Supportgroup] = N'MMIS Support Backend Apps Filenet' WHERE Servername = N'ah57sdcww00076' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet', [Purpose] = N'Filenet ICN Web 1', [Supportgroup] = N'MMIS Support Backend Apps Filenet' WHERE Servername = N'ah57sdcww00077' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet', [Purpose] = N'Filenet ICN Web 2', [Supportgroup] = N'MMIS Support Backend Apps Filenet' WHERE Servername = N'ah57sdcww00078' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet', [Purpose] = N'Filenet CPE / ICN Web 1', [Supportgroup] = N'MMIS Support Backend Apps Filenet' WHERE Servername = N'ah57sdcww00079' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet', [Purpose] = N'Filenet CPE / ICN Web 2', [Supportgroup] = N'MMIS Support Backend Apps Filenet' WHERE Servername = N'ah57sdcww00080' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet', [Purpose] = N'Filenet CPE / ICN Web 1', [Supportgroup] = N'MMIS Support Backend Apps Filenet' WHERE Servername = N'ah57sdcww00081' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet', [Purpose] = N'Filenet CPE / ICN Web 2', [Supportgroup] = N'MMIS Support Backend Apps Filenet' WHERE Servername = N'ah57sdcww00082' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LEGACY EPLS', [Purpose] = N'EPLS Prod server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww00086' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'LEGACY EPLS', [Purpose] = N'EPLS Test server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww00087' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Project WorkBook (PWB)', [Purpose] = N'OS Web / App', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww00088' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TPrime', [Purpose] = N'Web Front End', [Supportgroup] = N'TPRIME SW ADMIN' WHERE Servername = N'ah57sdcww00100' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TPrime', [Purpose] = N'Web Front End', [Supportgroup] = N'MMIS Change Backend Apps Sharepoint' WHERE Servername = N'ah57sdcww00101' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TPrime', [Purpose] = N'Web Front End', [Supportgroup] = N'MMIS Change Backend Apps Sharepoint' WHERE Servername = N'ah57sdcww00102' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'RRI - Resource Recognition Incorporated', [Purpose] = N'Navigator Web / App', [Supportgroup] = N'MMIS Support Backend Apps Formworks' WHERE Servername = N'ah57sdcww00103' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'RRI - Resource Recognition Incorporated', [Purpose] = N'WebKey 2 / App', [Supportgroup] = N'MMIS Support Backend Apps Formworks' WHERE Servername = N'ah57sdcww00104' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Legal Hold', [Purpose] = N'Legal Hold Web / App' WHERE Servername = N'ah57sdcww00112' AND ([System] IS NULL OR [Purpose] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Legal Hold', [Purpose] = N'Legal Hold Web / App', [Supportgroup] = N'AWS SW Admin' WHERE Servername = N'ah57sdcww00113' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'Appeon Build Server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww00114' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Decision Analytics', [Purpose] = N'App / WebApp / Win SVC / IIS', [Supportgroup] = N'MMIS Support Decision Analytics Dashboard' WHERE Servername = N'ah57sdcww00115' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Decision Analytics', [Purpose] = N'App / WebApp / Win SVC / IIS', [Supportgroup] = N'MMIS Support Decision Analytics Dashboard' WHERE Servername = N'ah57sdcww00116' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Decision Analytics', [Purpose] = N'App / WebApp / Win SVC/ IIS', [Supportgroup] = N'MMIS Support Decision Analytics Dashboard' WHERE Servername = N'ah57sdcww00117' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SharePoint UAT WEB', [Purpose] = N'SharePoint UAT WEB', [Supportgroup] = N'MMIS Change Backend Apps Sharepoint' WHERE Servername = N'ah57sdcww00118' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer Web', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww00119' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer Web', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww00120' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer Web', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww00121' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Appeon', [Purpose] = N'PowerServer Web', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww00122' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'EPLS', [Purpose] = N'EPLS Prod Server', [Supportgroup] = N'EPLS SW ADMIN' WHERE Servername = N'ah57sdcww01081' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Provider Web Service', [Purpose] = N'PWS Web / App', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww01082' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Provider Web Service', [Purpose] = N'PWS Web / App', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww01083' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SURS - Surveillance and Utilization Revi', [Purpose] = N'SURS Web / App', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww01084' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SURS - Surveillance and Utilization Revi', [Purpose] = N'SURS Web / App', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww01085' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TPM - Trading Partner Management', [Purpose] = N'EDI Trading Partner Web Server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww01086' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TPM - Trading Partner Management', [Purpose] = N'EDI Trading Partner Web Server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww01087' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'HIE - Health Information Exchange', [Purpose] = N'HIE Web Server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww01088' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'HIE - Health Information Exchange', [Purpose] = N'HIE Web Server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww01089' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TMSIS - Transformed Medicaid Statistical', [Purpose] = N'TMSIS Web', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww01090' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TMSIS - Transformed Medicaid Statistical', [Purpose] = N'TMSIS Web', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww01091' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Exstream', [Purpose] = N'License Server', [Supportgroup] = N'EXSTREAM SW ADMIN' WHERE Servername = N'ah57sdcww01150' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Exstream', [Purpose] = N'Web/App Server', [Supportgroup] = N'EXSTREAM SW ADMIN' WHERE Servername = N'ah57sdcww01151' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS', [Purpose] = N'Web Server', [Supportgroup] = N'TCOS SW ADMIN' WHERE Servername = N'ah57sdcww01152' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS', [Purpose] = N'Web Server', [Supportgroup] = N'TCOS SW ADMIN' WHERE Servername = N'ah57sdcww01153' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS', [Purpose] = N'Web Server', [Supportgroup] = N'TCOS SW ADMIN' WHERE Servername = N'ah57sdcww01154' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Visual Application', [Purpose] = N'EDI Trading Partner Web Server', [Supportgroup] = N'VISUAL APP SW ADMIN' WHERE Servername = N'ah57sdcww01155' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Visual Application', [Purpose] = N'SURS Web Server', [Supportgroup] = N'VISUAL APP SW ADMIN' WHERE Servername = N'ah57sdcww01156' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Visual Application', [Purpose] = N'HIE Web Server', [Supportgroup] = N'VISUAL APP SW ADMIN' WHERE Servername = N'ah57sdcww01157' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Visual Application', [Purpose] = N'TMSIS_CMS Web Server', [Supportgroup] = N'VISUAL APP SW ADMIN' WHERE Servername = N'ah57sdcww01158' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Online Systems', [Purpose] = N'OS Web / App', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww01165' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Online Systems', [Purpose] = N'OS Web / App', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww01166' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Online Systems', [Purpose] = N'OS Web / App', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww01167' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Doctors Lookup', [Purpose] = N'Doctor Lookup Web server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww01168' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Doctors Lookup', [Purpose] = N'Doctor Lookup Web server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww01169' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Doctors Lookup', [Purpose] = N'Doctor Lookup Web server', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww01170' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Online Systems', [Purpose] = N'OS Web / App', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww01171' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Online Systems', [Purpose] = N'OS Web / App', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww01172' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Online Systems', [Purpose] = N'OS Web / App', [Supportgroup] = N'MMIS Support Visual Apps Visual Apps (Appeon/TCOS/PWB/.NET)' WHERE Servername = N'ah57sdcww01173' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM', [Purpose] = N'ALM Web Node 1', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcww20002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ALM', [Purpose] = N'ALM Web Node 2', [Supportgroup] = N'ALM SW ADMIN' WHERE Servername = N'ah57sdcww20003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Jira for MMIS', [Purpose] = N'Jira Web Server', [Supportgroup] = N'AGM SW ADMIN' WHERE Servername = N'ah57sdcww20004' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TPrime SharePoint WFE', [Purpose] = N'SharePoint SSRS/Power Pivot Front End Server', [Supportgroup] = N'MMIS Support Backend Apps SharePoint' WHERE Servername = N'ah57sdcwx00001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TPrime', [Purpose] = N'SharePoint SSRS/Power Pivot', [Supportgroup] = N'MMIS Change Backend Apps Sharepoint' WHERE Servername = N'ah57sdcwx00002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TPrime', [Purpose] = N'T-Prime SharePoint SSRS/Power Pivot Front End Server', [Supportgroup] = N'TPRIME SW ADMIN' WHERE Servername = N'ah57sdcwx00003' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'RRI - Resource Recognition Incorporated', [Purpose] = N'Dev - All', [Supportgroup] = N'MMIS Support Backend Apps Formworks' WHERE Servername = N'ah57sdcwx00004' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Decision Analytics', [Purpose] = N'UAT / STAGING Power BI, SSRS, & SSAS', [Supportgroup] = N'WINDOWS OS ADMIN STS' WHERE Servername = N'ah57sdcwx00007' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Decision Analytics', [Purpose] = N'DEV Power BI, SSRS, & SSAS', [Supportgroup] = N'WINDOWS OS ADMIN STS' WHERE Servername = N'ah57sdcwx00008' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Edifecs', [Purpose] = N'Edifecs - soon to be decommissioned' WHERE Servername = N'dc5719052wa015' AND ([System] IS NULL OR [Purpose] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Edifecs', [Purpose] = N'Edifecs - soon to be decommissioned' WHERE Servername = N'dc5719052wa016' AND ([System] IS NULL OR [Purpose] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Edifecs', [Purpose] = N'Edifecs - soon to be decommissioned' WHERE Servername = N'dc5719052wa017' AND ([System] IS NULL OR [Purpose] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ProLaw', [Purpose] = N'ProLaw', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719052waa221' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'IBM Rational License Server', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719052wai102' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Edifecs', [Purpose] = N'Edifecs - soon to be decommissioned' WHERE Servername = N'dc5719052wb001' AND ([System] IS NULL OR [Purpose] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Edifecs', [Purpose] = N'Edifecs - soon to be decommissioned' WHERE Servername = N'dc5719052wb002' AND ([System] IS NULL OR [Purpose] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'State File Share Seibel Archive State Security Archive', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719052wb100' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Edifecs', [Purpose] = N'EDIFECS Dev' WHERE Servername = N'dc5719??' AND ([System] IS NULL OR [Purpose] IS NULL);
UPDATE #TempSnowServer SET [System] = N'GWT Infrastructure', [Purpose] = N'File Server', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719cfs60' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'GWT Infrastructure', [Purpose] = N'Dynatrace', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719dy249' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'DynaTrace', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719dyt250' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'iC DEV', [Purpose] = N'New home for fat client builds for testing', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719icdev103' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Pharmacy Archive', [Purpose] = N'File Server', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719pa188' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'POC for PXE Boot Server (Server 2019 Datacenter)', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719pxe32' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [Purpose] = N'RRI' WHERE Servername = N'dc5719rri67' AND ([Purpose] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'Form Works 5 RRI Bridges/WebKey (Server 2019 Datacenter)', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719rri68' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'Form Works 5 RRI Bridges/WebKey (Server 2019 Datacenter)', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719rri69' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'Power BI POC server', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719swap99' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'VCenter Appliance', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719vctr245' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'TEDS Share', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wa107' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'File Server', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719wa13' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'3Par Reports Server', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wa131' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'RightFax', [Purpose] = N'RightFax', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wa171' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'RightFax', [Purpose] = N'RightFax', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wa172' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'RightFax', [Purpose] = N'RightFax', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wa173' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'RightFax', [Purpose] = N'RightFax', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wa174' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'RightFax', [Purpose] = N'RightFax', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wa175' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'TN ANYTIME DEV SERVER, VA Sandbox' WHERE Servername = N'dc5719wa35' AND ([System] IS NULL OR [Purpose] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'SQL 2016 temporary jump server for SharePoint migration to 2019', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wa75' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'SQL 2016 temporary jump server for SharePoint migration to 2019', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wa76' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ProLaw', [Purpose] = N'Prolaw SSRS Server', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719waa193' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SharePoint', [Purpose] = N'Sharepoint', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719waa66' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719waa73' AND ([Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719waa74' AND ([Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'NET 2019 VM Template', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719waf79' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'Domain Controller', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap10' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'IWI', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap100' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'Teammate IIS', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap101' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'File Server', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap102' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'Teamate', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap105' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'HCI File Server', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap106' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'Prod SSRS for TEAMS', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719wap108' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ProLaw', [Purpose] = N'ProLaw, File repository for BIDS, TSFSqlServer, Kepro application', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719wap109' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Domain Controller - TENNCARE - Primary', [Purpose] = N'Domain Controller - TENNCARE - Primary', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap11' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'Requested by Tina Rowe DXC DBA for remote Oracle Installs.', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap115' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'Was DC5719WAP99', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap116' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Domain Controller - TENNCARE - Tertiary', [Purpose] = N'Domain Controller - TENNCARE - Tertiary', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap12' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'PW Reset .Net solution', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap125' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'One of the State File Share Servers TeamMate?', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap127' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'One of the State File Share Servers OPI Share Server', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap128' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'State Eligibility Initiative File Share Server', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap129' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'MPI APP Server Requested by Clay' WHERE Servername = N'dc5719wap141' AND ([System] IS NULL OR [Purpose] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD', [Purpose] = N'CMOD', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap143' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet', [Purpose] = N'Filenet', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap144' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet', [Purpose] = N'Filenet', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap145' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FileNet', [Purpose] = N'Filenet', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap146' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD', [Purpose] = N'CMOD', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap147' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'CMOD', [Purpose] = N'CMOD', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap148' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap15' AND ([Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'Exstream Letters Server', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap150' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'Exstream Test Server', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap151' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'SQL CLUSTER Node 3', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719wap168h' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'AbacusLaw Software', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap170' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'RRI Deployment Server' WHERE Servername = N'dc5719wap19' AND ([System] IS NULL OR [Purpose] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Sparx', [Purpose] = N'Sparx', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap190' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Siebel', [Purpose] = N'Oracle DB Server/Siebel' WHERE Servername = N'dc5719wap196' AND ([System] IS NULL OR [Purpose] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'DEV .Net Solutions (Web Services & Console Apps)', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap34' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'Server 2019 Datacenter VM template. Used for builds and compliance scans', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap35' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'NET Domain Terminal License Server', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719wap39' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'IIS Server (Password reset tool)', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap42' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'HP FrontRange' WHERE Servername = N'dc5719wap50' AND ([System] IS NULL OR [Purpose] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'Tenncare DC SiteScope Server' WHERE Servername = N'dc5719wap56' AND ([System] IS NULL OR [Purpose] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'New Print server for Jetadmin replaced .51', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap57' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TPrime', [Purpose] = N'T-Prime Dev', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719wap60' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SharePoint', [Purpose] = N'Sharepoint', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719wap61' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Decision Analytics', [Purpose] = N'Decision Analytics', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719wap65' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SharePoint', [Purpose] = N'Sharepoint', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719wap67' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SharePoint', [Purpose] = N'SharePoint 2012 Cluster Host 1', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719wap68' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SharePoint', [Purpose] = N'SharePoint 2012 Cluster Host 2', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719wap69' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FIS - Formworks 5', [Purpose] = N'FIS - Formworks 5', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap70' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FIS - Formworks 5', [Purpose] = N'FIS - Formworks 5', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap71' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FIS - Formworks 5', [Purpose] = N'Application Server', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap72' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FIS - Formworks 5', [Purpose] = N'FIS - Formworks 5', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap73' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FIS - Formworks 5', [Purpose] = N'FIS - Formworks 5', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap74' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FIS - Formworks 5', [Purpose] = N'FIS - Formworks 5', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap75' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FIS - Formworks 5', [Purpose] = N'FIS - Formworks 5', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap76' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'FIS - Formworks 5', [Purpose] = N'FIS - Formworks 5', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap77' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'DEV .Net Solutions (Web Services & Console Apps)', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wap97' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'2012R2 Oracle Data Modeler/PAL', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wat119' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'ProLaw', [Purpose] = N'ProLaw non-prod', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wb02' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'SharePoint', [Purpose] = N'Sharepoint DEV SQL', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719wb62' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'SQL Cluster', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719wbb161' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'SQL Cluster', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719wbb162' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'File Server (SSAS)', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719wbb164' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'File Server (SSIS)', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719wbb165' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'SQL Server', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719wbb166' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'VMWare Application Extension server VMWare Update Manager Server VMWare Download Server Brocade Oneview SAN Manager Server Oneview vCenter Extension Server', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wbb30' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'Tenncare Shared SQL Server', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wdb122' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'Pharmacy Archive HDD Export Server', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wfs10' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'Production machine for Checkpoint and Ghost.' WHERE Servername = N'dc5719wl01313-t' AND ([System] IS NULL OR [Purpose] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'State Share', [Supportgroup] = N'MMIS Support INF Windows' WHERE Servername = N'dc5719wss180' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'TennCare Intranet', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wwa110' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'PWB NET Webserver' WHERE Servername = N'dc5719wwa111' AND ([System] IS NULL OR [Purpose] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'DSi Hold Dev Server', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wwa155' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'DSi RGC PROD Server', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wwa156' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wwa157' AND ([Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS - TN Anytime', [Purpose] = N'TCOS - TN Anytime', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wwp210' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS - TN Anytime', [Purpose] = N'TCOS - TN Anytime', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wwp211' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS - TN Anytime', [Purpose] = N'TCOS - TN Anytime', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wwp212' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS - TN Anytime', [Purpose] = N'TCOS - TN Anytime', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wwp213' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS - TN Anytime', [Purpose] = N'TCOS - TN Anytime', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wwp214' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS - TN Anytime', [Purpose] = N'TCOS - TN Anytime', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wwp215' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS - TN Anytime', [Purpose] = N'TCOS - TN Anytime', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wwp216' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'MMIS', [Purpose] = N'TCOS', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wwp224' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS - TN Anytime', [Purpose] = N'TCOS - TN Anytime', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wwp40' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS - TN Anytime', [Purpose] = N'TCOS - TN Anytime', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wwp41' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS - TN Anytime', [Purpose] = N'TCOS - TN Anytime', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wwp43' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS - TN Anytime', [Purpose] = N'TCOS - TN Anytime', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wwp45' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS - TN Anytime', [Purpose] = N'TCOS - TN Anytime', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wwp46' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS - TN Anytime', [Purpose] = N'TCOS - TN Anytime', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wwp47' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'TCOS - TN Anytime', [Purpose] = N'TCOS - TN Anytime', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wwp48' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'IIS (StaffMap) SQL Server', [Supportgroup] = N'Vendor IT Support' WHERE Servername = N'dc5719wws140' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'TennCare Call Recording' WHERE Servername = N'sgprd47calbas01' AND ([System] IS NULL OR [Purpose] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'CrowdStrike', [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'tennacccwdstk' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'tnxixcpctr1' AND ([Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'tnxixcpctr2' AND ([Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'tnxixcpmgm01' AND ([Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Security', [Purpose] = N'Azure Analytics', [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'tnxixe1warc001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Security', [Purpose] = N'Azure Analytics', [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'tnxixe1wccs001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'VM AD Domain Controller', [Purpose] = N'Azure Analytics', [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'tnxixe1wd001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'VM AD Backup Domain Controller', [Purpose] = N'Azure Analytics', [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'tnxixe1wd002' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'VM AD Mgt. Server (Restricted)', [Purpose] = N'Azure Analytics', [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'tnxixe1wmgm001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'Security', [Supportgroup] = N'Azure SW Admin GWT' WHERE Servername = N'tnxixe1wqly001' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'RHEL 8 EOC migration dev', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmdeveueoc01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'SAS Server (DEV)', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmdeveusas01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'SAS Server (DEV)', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmdeveusas02' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'SAS Server (DEV)', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmdeveusas03' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'SAS Server (DEV)', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmdeveusas04' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'Tableau development server', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmdeveutab01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'admin workstation', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeuadmin01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'SAS Server (DEV)', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeuans01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'CIS CAT Pro security compliance workstation', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeucat01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'Azure Domain Controller', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeudc02' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'Azure on-premises data gateway', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeudgw01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'Azure on-premises data gateway', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeudgw02' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'data science PROD runtime host', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeuds01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'EOC Server', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeueoc01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'EOC Server', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeueoc02' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'EOC Server', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeueoc03' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'Jumpbox server for firewall management', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeujump01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'Firewall virtual machine', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeupa02' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'Firewall virtual machine', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeupa03' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'Firewall virtual machine', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeupa04' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeupa05' AND ([System] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'terminal services management host', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeurds01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'terminal services PROD session host', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeurds10' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'SAS Server (PRD)', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeusas01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'SAS Server (PRD)', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeusas02' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'SAS Server (PRD)', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeusas03' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'SAS Server (PRD)', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeusas04' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'Spunk Heavy Forwarder', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeuspk01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'Spunk Heavy Forwarder', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeuspk02' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'Spunk Distribution Server', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeuspk03' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'Splunk heavy forwarder managed by Gainwell', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeuspk04' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'TripWire Server', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeustw01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Tableau (EOC)', [Purpose] = N'Azure Tableau (EOC)', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmprdeutab01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'admin server in sandbox environment', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmsbxeuadmin01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'data science SAND runtime host', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmsbxeuds01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'data science UAT runtime host', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmuateuds01' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'Azure Analytics', [Purpose] = N'terminal services UAT session host', [Supportgroup] = N'MMIS Support Data Warehouse Azure' WHERE Servername = N'vmuateurds10' AND ([System] IS NULL OR [Purpose] IS NULL OR [Supportgroup] IS NULL);
UPDATE #TempSnowServer SET [System] = N'DXC Infrastructure', [Purpose] = N'VM Templates' WHERE Servername = N'win-2ijjqtt29ru' AND ([System] IS NULL OR [Purpose] IS NULL);


--6. step 6 -- update script the merge to from temp snow table to  temp sts.server, using merge.sql
-- ============================================================================
-- Same MERGE as merge.sql, retargeted at the step 4 temp tables (#TempServer/#TempSnowServer
-- are session-scoped copies, not the real sts.server/ita.snowserver) so it can be dry-run safely.
-- Dedup logic (Rn = 1, newest Sysupdatedon wins) and the full COALESCE-overwrite column list are
-- unchanged from merge.sql.
MERGE  	into #TempServer AS TR
   USING    ( SELECT distinct LTRIM(RTRIM( [ita].[fnSnowServer](Servername,'Servername'))) AS Servername
			  ,Servername As OldServername
			  ,serverid
			  ,changeordernumber
			  ,ipaddress
			 ,operationalstatus
			 ,System
			 ,Process
			  ,Environment
			 ,location
			  ,Sysclassname
			 ,Classification
			 ,os
			 ,osdomain
			 ,company
			 ,ram
			 ,dsnDomain
			 ,OSversion
			 ,supportedby
			 ,managedby
			 ,rn
			  ,Sysupdatedon
			  ,syscreatedon
    FROM     (
	        -----retrive server information which have maximim Sysupdatedon
             SELECT		i.*, ROW_NUMBER() OVER (PARTITION BY LTRIM(RTRIM( [ita].[fnSnowServer](i.Servername,'Servername')))
		     ORDER BY   Sysupdatedon  DESC,LTRIM(RTRIM( [ita].[fnSnowServer](i.Servername,'Servername')))) AS Rn
		     FROM       #TempSnowServer i
		     where  	 lTRIM(RTRIM( [ita].[fnSnowServer] (Servername,'Servername'))) !=''
			 or			 lTRIM(RTRIM( [ita].[fnSnowServer] (Servername,'Servername'))) is not null
             ) r
    WHERE    Rn = 1
			 )  AS t
			 On T.Servername   = [ita].[fnSnowServer](LTRIM(RTRIM(TR.Servername)),'Servername')

    WHEN   MATCHED
    -- Server name change order number, ENV, Operation status, and classification
    THEN	  UPDATE
    Set       tr.IPAddress= COALESCE(LTRIM(RTRIM(t.ipaddress)) ,tr.IPAddress)
	          ,tr.ChangeOrderNumber= COALESCE(LTRIM(RTRIM(t.ChangeorderNumber)) ,tr.ChangeOrderNumber)
			  ,tr.operationalstatus= COALESCE(LTRIM(RTRIM(t.operationalstatus)) ,tr.operationalstatus)
			  ,tr.system= COALESCE(LTRIM(RTRIM(t.Process)) ,tr.system)
			  ,tr.primarysubsystem= COALESCE(LTRIM(RTRIM(t.environment)) ,tr.primarysubsystem)
			  ,tr.ENV= COALESCE(LTRIM(RTRIM(t.classification)) ,tr.ENV)
			  ,tr.Operatingsystem= COALESCE(LTRIM(RTRIM(t.os)) ,tr.Operatingsystem)
			  ,tr.Businessowner= COALESCE(LTRIM(RTRIM(t.managedby)) ,tr.Businessowner)
			  ,tr.ServerDescription = COALESCE(LTRIM(RTRIM(t.sysclassname)) ,tr.ServerDescription)
	          ,tr.DataCenter= COALESCE(LTRIM(RTRIM(t.location)) ,tr.DataCenter)
		     , tr.AssetOwner= COALESCE(LTRIM(RTRIM(t.osdomain)) ,tr.AssetOwner)
    		,  tr.Agency = COALESCE(LTRIM(RTRIM(t.company)) ,tr.Agency )
			  ,tr.memory= COALESCE(LTRIM(RTRIM(t.ram)) ,tr.memory)
	          ,tr.LocationExtension= COALESCE(LTRIM(RTRIM(t.dsnDomain)) ,tr.LocationExtension)
			  ,tr.OSversion= COALESCE(LTRIM(RTRIM(t.OSversion)) ,tr.OSversion)
              ,tr.TechnicalOwner= COALESCE(LTRIM(RTRIM(t.supportedby)) ,tr.TechnicalOwner)
	    ---if its source is null
               ,tr.Source = COALESCE( tr.Source,'SNOWServer')
			   --if it is matched update the primary source'SnowServer'
			   ,tr.primarySource= 'SNOWServer'
	           ,tr.BillingSource=COALESCE(tr.BillingSource, 'SNOWServer')
			   ,tr.BillingType='On-Prem'
			   ,tr.SnowServerFlag=1
			    ,tr.Decom=0
			  ; 


