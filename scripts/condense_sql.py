#!/usr/bin/env python3
"""
Condense sts_server_comparison.sql to put 7 servers per line
"""

input_file = r"c:\projects\GainwellServers_20260820\output\sts_server_comparison.sql"
output_file = r"c:\projects\GainwellServers_20260820\output\sts_server_comparison.sql"

# Read the file
with open(input_file, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Find the section with server names
output_lines = []
server_lines = []
in_server_section = False
servers_to_group = []

for line in lines:
    # Check if this line contains a server name (starts with --   N')
    if line.strip().startswith("--   N'"):
        in_server_section = True
        # Extract the server name
        match = line.strip()
        servers_to_group.append(match)
    elif in_server_section and line.strip() == "-- )":
        # End of server section
        in_server_section = False
        
        # Group servers 7 per line
        for i in range(0, len(servers_to_group), 7):
            group = servers_to_group[i:i+7]
            grouped_line = "-- " + " ".join(group) + "\n"
            output_lines.append(grouped_line)
        
        servers_to_group = []
        output_lines.append(line)
    elif in_server_section:
        # We've hit something else while in server section
        in_server_section = False
        
        # Group remaining servers
        if servers_to_group:
            for i in range(0, len(servers_to_group), 7):
                group = servers_to_group[i:i+7]
                grouped_line = "-- " + " ".join(group) + "\n"
                output_lines.append(grouped_line)
        
        servers_to_group = []
        output_lines.append(line)
    else:
        output_lines.append(line)

# Write the output
with open(output_file, 'w', encoding='utf-8') as f:
    f.writelines(output_lines)

print(f"✓ Condensed {len(servers_to_group)} servers to 7 per line")
print(f"✓ File updated: {output_file}")
