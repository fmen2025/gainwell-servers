# 

## db table
the Staging table **table** name is `ita.snowserver` 

## source file 
read data tab from Excel file that located on  `C:\Feven\Hanna\2026\Gainwell Servers_20260820.xlsx`

## secrets
Never expose my Secrets in the chat window or commit/push to git remote. Also, only use the parameter name for secrets not the values in Python/PowerShell Scripts. Read the secrets from my .env file.

**Reminder**: Never print, echo, or display the contents/values of my `.env` file in the chat window, terminal output, logs, or any other place. Never read the `.env` file directly to show its contents. Always access secrets only through environment variable names (e.g. `os.environ["DB_CONN"]`) inside scripts — never hardcode or restate the actual values anywhere.

## Input Files:

mapping.md containing the source-to-staging column mappings.
Source Excel/reading file containing server information.

## Process:

Review mapping.md and identify how each source column maps to the database columns.
Compare the source file records against the existing database records using ServerName as the primary key.
ServerName matching is case-insensitive (ita.snowserver.Servername uses SQL_Latin1_General_CP1_CI_AS
collation), so comparisons must not treat differently-cased ServerNames as different servers.
ServerNames that are duplicated in the source file and/or in the database are excluded from the
new/update recommendations (not usable as a comparison key) and are listed separately for human review.
Identify servers that do not exist in the database.
Identify servers that exist but have NULL values in mapped fields.
For each missing server, prepare an insert recommendation.
For existing servers with NULL fields, identify which values can be populated from the source file.
Do not recommend changes to fields that already contain non-NULL values.

**Read-only constraint**: This is a comparison/reporting task only. Never INSERT, UPDATE, DELETE, or otherwise modify the database. Only produce recommendation reports using .sql script and file for (new servers to insert, existing servers/fields to update) for human review — do not execute any of those changes against `ita.snowserver`.


## Expected Output:

List of new servers not found in the database.
List of existing servers with NULL fields that can be updated.
Mapping evidence from mapping.md showing how source columns correspond to database columns.

Review mapping.md and validate the source-to-database column mappings using actual repository evidence.

Compare source file records with existing database records. Identify:
1. Servers that do not exist in the database.
2. Servers that exist but have NULL values that can be populated from the source file.

Do not assume mappings, keys, tables, columns, or business rules. Do not hallucinate or guess.

If any mapping, comparison key, or requirement is unclear, stop and ask a specific question before proceeding.

## Rule for script

It is acceptable to use Python for analysis if needed, but the final solution must be .SQL-based.
Generate the SQL required to:
- Identify missing servers.
- Identify existing servers with NULL values.

**Recommendation script format**: The `.sql` output must contain concrete, ready-to-review statements a
human can act on — not only SELECT reports:
- One `INSERT INTO ita.snowserver (...) VALUES (...)` statement per missing server.
- One `UPDATE ita.snowserver SET [field] = '...' WHERE Servername = '...' AND [field] IS NULL` statement
  per NULL field that can be populated.
All INSERT/UPDATE statements must be written as SQL comments (commented out) so the file can never be
run as-is and accidentally change the database — a human must deliberately review and uncomment the
specific statements they choose to apply. Do not execute any INSERT/UPDATE/DELETE from this workflow.

## logs
 update the steps that done on log file `log.md`