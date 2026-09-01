#!/usr/bin/env python3
"""
Consolidate 2824 individual UPDATE statements into 946 per-server consolidated statements
"""
import re
from collections import defaultdict

# Read the SQL file
input_file = r"c:\projects\GainwellServers_20260820\output\compare_servers.sql"
output_file = r"c:\projects\GainwellServers_20260820\output\compare_servers_consolidated.sql"

with open(input_file, 'r', encoding='utf-8') as f:
    content = f.read()

# Find all header/summary lines
header_end = content.find("-- 2) RECOMMENDED UPDATEs for NULL fields")
header = content[:header_end]

# Extract all UPDATE statements
update_pattern = r"-- UPDATE ita\.snowserver SET \[([^\]]+)\] = (N'[^']*(?:'[^']*')*[^']*') WHERE Servername = N'([^']+)' AND \[[^\]]+\] IS NULL;"
matches = re.findall(update_pattern, content)

# Group by ServerName
grouped = defaultdict(list)
for field, value, servername in matches:
    grouped[servername].append((field, value))

print(f"Total distinct servers: {len(grouped)}")
print(f"Total field updates: {len(matches)}")

# Create consolidated SQL statements
consolidated_statements = []
for servername in sorted(grouped.keys()):
    fields = grouped[servername]
    
    # Build SET clause
    set_clauses = [f"[{field}] = {value}" for field, value in fields]
    set_clause = ", ".join(set_clauses)
    
    # Build WHERE clause
    field_names = [field for field, _ in fields]
    where_conditions = " OR ".join([f"[{f}] IS NULL" for f in field_names])
    
    # Create full statement
    statement = f"-- UPDATE ita.snowserver SET {set_clause} WHERE Servername = N'{servername}' AND ({where_conditions});"
    consolidated_statements.append(statement)

# Write consolidated file
with open(output_file, 'w', encoding='utf-8') as f:
    f.write(header)
    f.write("\n-- 2) RECOMMENDED UPDATEs for NULL fields (CONSOLIDATED - 1 UPDATE per server)\n")
    f.write(f"-- Total consolidated updates: {len(consolidated_statements)}\n\n")
    
    for stmt in consolidated_statements:
        f.write(stmt + "\n")

print(f"\nConsolidated file created: {output_file}")
print(f"Original statements: {len(matches)}")
print(f"Consolidated statements: {len(consolidated_statements)}")
print(f"Reduction: {len(matches) - len(consolidated_statements)} fewer statements")
