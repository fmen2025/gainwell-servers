#!/usr/bin/env python3
"""
Remove line breaks within SQL UPDATE statements
"""
import re

files = [
    r"c:\projects\GainwellServers_20260820\output\compare_servers.sql",
    r"c:\projects\GainwellServers_20260820\output\compare_servers_consolidated.sql"
]

for filepath in files:
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Find lines that start with "-- UPDATE" or continue them (not starting with --)
        lines = content.split('\n')
        fixed_lines = []
        i = 0
        
        while i < len(lines):
            line = lines[i]
            
            # If this is an UPDATE statement
            if line.strip().startswith('-- UPDATE'):
                # Collect this line and any continuation lines
                full_statement = line
                i += 1
                
                # Continue collecting lines until we find the closing semicolon
                while i < len(lines) and not lines[i].strip().endswith(';'):
                    next_line = lines[i].strip()
                    if next_line:  # Skip empty lines
                        # Replace newlines with space in the accumulated statement
                        full_statement += ' ' + next_line
                    i += 1
                
                # Add the final line with semicolon
                if i < len(lines):
                    full_statement += ' ' + lines[i].strip()
                    i += 1
                
                fixed_lines.append(full_statement)
            else:
                fixed_lines.append(line)
                i += 1
        
        # Write back
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write('\n'.join(fixed_lines))
        
        print(f"✓ Fixed: {filepath}")
    
    except FileNotFoundError:
        print(f"⚠ File not found: {filepath}")
    except Exception as e:
        print(f"✗ Error processing {filepath}: {e}")
