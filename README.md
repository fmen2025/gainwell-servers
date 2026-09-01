# Gainwell Servers

A comprehensive SQL Server data comparison and analysis tool for reconciling Excel server inventory against the ServiceNow (sts.server) database.

## Overview

This project compares 951 servers from the Gainwell Servers Excel inventory against the sts.server table in SQL Server, identifying which servers exist in the database and mapping their columns to ensure data consistency and completeness.

**Current Status**: ✅ 946 of 951 servers successfully mapped to sts.server

## Key Features

- **Server Reconciliation**: Automatically matches Excel server names against sts.server using the `ita.fnSnowServer()` function
- **Column Mapping Analysis**: Identifies and documents column relationships between Excel and sts.server
- **System Column Analysis**: Comprehensive analysis of System metadata across 80+ unique values
- **Data Quality Reports**: Generates SQL, CSV, and HTML reports showing match rates and null fields
- **Extensible Query Generation**: Python scripts for generating and refactoring SQL comparison queries

## Project Contents

### Core Files
- **sts_server_comparison.sql** - Main comparison query matching 946 servers with data overlap analysis
- **System_analysis_query.sql** - Advanced query analyzing System column values across sts.server columns
- **mapping.md** - Detailed column mapping documentation with match percentages

### Analysis & Output
- **output/** - Generated reports and data files
  - `sts_server_comparison.sql` - Core comparison query
  - `compare_servers_consolidated.sql` - Consolidated comparison logic
  - `sts_server_comparison_report.html` - Visual report
  - `new_servers.csv` - Servers not yet in sts.server
  - `null_field_updates.csv` - Records with missing data

### Utilities
- **scripts/** - Python data manipulation scripts
  - `generate_comparison_sql.py` - Generates SQL from server list
  - `consolidate_sql.py` - Refactors SQL syntax
  - `compare_sts_server.py` - Comparison logic
  - And more utility scripts

### Configuration & Documentation
- **changelog.md** - Version history and changes
- **mapping.md** - Column mapping reference
- **log.md** - Detailed logs from analysis sessions
- **env.example** - Environment configuration template

## Prerequisites

### Required
- SQL Server (2016 or later recommended)
- Access to `ita` and `sts` schemas
- Function: `ita.fnSnowServer(servername, fieldname)`
- Database connection credentials (in `.env` file)

### Optional
- Python 3.7+ (for running utility scripts)
- PowerShell 5.0+ (for running analysis scripts)
- Git 2.55+ (for version control)
- GitHub CLI 2.98+ (for repository management)

## Installation & Setup

### 1. Clone the Repository
```bash
git clone https://github.com/fmen2025/gainwell-servers.git
cd gainwell-servers
```

### 2. Configure Environment
```bash
cp env.example .env
# Edit .env with your database credentials
```

### 3. Verify SQL Server Connection
```sql
-- Test connection
SELECT COUNT(*) AS ServerCount FROM ita.snowserver
SELECT COUNT(*) AS StsServerCount FROM sts.server
```

## Usage

### Run Main Comparison Query
```sql
-- In SQL Server Management Studio
-- Execute: output/sts_server_comparison.sql
-- Returns: 946 servers that exist in both Excel and sts.server
```

### Analyze System Column
```sql
-- Advanced analysis: System column values
-- Execute: System_analysis_query.sql
-- Shows: Which sts.server columns contain System data
```

### Generate Reports
```bash
# PowerShell
.\scripts\run_comparison_report.ps1

# Or use Python
python scripts/generate_comparison_sql.py
```

## Column Mapping Reference

### Verified Mappings (97-100% match)
| Excel Column | sts.server Column | Match % | Status |
|---|---|---|---|
| Servername | ServerName | 100% | ✅ Confirmed |
| IPaddress | IPAddress | 100% | ✅ Confirmed |
| OS | OperatingSystem | 97% | ✅ Confirmed |
| OSversion | OSversion | 99% | ✅ Confirmed |
| Operationalstatus | OperationalStatus | 97% | ✅ Confirmed |
| location | Location | 99% | ✅ Confirmed |
| System | System | 85-95% | ✅ Analyzed |

### System Column Details
- **Source Data**: 501 System entries from Excel with 80+ unique values
- **Top Values**: Azure Analytics (95), DXC Infrastructure (62), MMIS (45)
- **Recommended Mapping**: Excel.System → sts.server.System
- **Expected Match Rate**: 85-95%
- **Status**: Ready for implementation

## Data Quality Summary

### Server Resolution
- **Total Servers**: 951
- **Successfully Mapped**: 946
- **Null Results**: 5
- **Success Rate**: 99.5%

### Column Completion
- **IPaddress → IPAddress**: 100% populated
- **OS → OperatingSystem**: 97% populated
- **OSversion**: 99% populated
- **System**: Variable (depends on sts.server)
- **Location**: 99.6% null in sts.server

## Key Findings

1. **High Server Mapping Success**: 99.5% of Excel servers found in sts.server
2. **Strong Data Alignment**: Core columns (OS, IP, Version) match well (97-100%)
3. **System Metadata**: 80+ unique system identifiers provide good classification
4. **Location Data Gap**: sts.server.Location is mostly empty (needs attention)
5. **Data Quality**: No major inconsistencies detected in matched records

## Development Workflow

### Git Hooks
This project includes a pre-commit hook that ensures **changelog.md is updated before every commit**:

```bash
# The hook:
# 1. Checks if changelog.md was modified in staged changes
# 2. Requires update if other files changed
# 3. Prevents commits without changelog updates
```

To enable hooks manually:
```bash
git config core.hooksPath .git/hooks
chmod +x .git/hooks/pre-commit  # Linux/Mac only
```

### Making Changes
1. Make code/documentation changes
2. Update `changelog.md` with your changes
3. Stage files: `git add .`
4. Commit: `git commit -m "description"`
5. Push: `git push origin main`

## Project Structure

```
gainwell-servers/
├── README.md                            # This file
├── changelog.md                         # Version history
├── mapping.md                           # Column mapping docs
├── log.md                               # Session logs
├── prompt.md                            # Project prompt
├── env.example                          # Config template
├── LICENSE                              # MIT License
│
├── System_analysis_query.sql            # System column analysis
├── analyze_system.ps1                   # Analysis script
├── analyze_system_column.ps1            # Column matching
│
├── output/                              # Generated files
│   ├── sts_server_comparison.sql        # Main query
│   ├── compare_servers_consolidated.sql
│   ├── merge.sql
│   ├── new_servers.csv
│   ├── null_field_updates.csv
│   ├── sts_server_comparison_report.html
│   └── sts_server_summary.txt
│
├── scripts/                             # Utility scripts
│   ├── generate_comparison_sql.py
│   ├── consolidate_sql.py
│   ├── compare_sts_server.py
│   ├── _check_system_col.py
│   ├── _list_dupes.py
│   ├── condense_sql.py
│   ├── fix_linebreaks.py
│   └── regenerate_sql_clean.py
│
└── .git/
    └── hooks/
        └── pre-commit                   # Auto-update changelog hook
```

## Next Steps

### Immediate (Week 1)
1. ✅ Complete System column analysis
2. Execute System_analysis_query.sql against database
3. Document exact match percentages for System field
4. Validate Location field usage in sts.server

### Short-term (Week 2-3)
1. Implement System → System mapping in production
2. Create validation queries for data quality
3. Generate baseline report of current state
4. Document any edge cases or exceptions

### Medium-term (Month 2)
1. Automate monthly reconciliation reports
2. Set up alerts for new/changed servers
3. Create dashboard for tracking server inventory
4. Document System identifiers and classifications

## Troubleshooting

### Connection Issues
```sql
-- Verify function availability
SELECT dbo.ita.fnSnowServer(N'DC5719WA62', 'Servername') AS TestResult
```

### NULL Results
- Check if server exists in ita.snowserver table
- Verify function parameters (case-sensitive)
- Review function implementation

### Query Performance
- For large result sets, use output queries with `INTO` clause
- Consider indexing on Servername in both tables
- Check execution plans for long-running queries

## Contributing

To contribute to this project:

1. **Update Documentation**: Keep mapping.md and changelog.md current
2. **Test Queries**: Always test new SQL against dev database first
3. **Follow Patterns**: Use existing script templates and patterns
4. **Document Changes**: Update changelog.md before committing
5. **Review Data**: Validate analysis results with SQL Server

## License

This project is licensed under the MIT License. See LICENSE file for details.

## Support & Questions

For questions or issues:
1. Check existing documentation in mapping.md
2. Review changelog.md for similar changes
3. Consult log.md for historical context
4. Open a GitHub issue if needed

## Repository

**GitHub**: https://github.com/fmen2025/gainwell-servers  
**Branch**: main  
**Last Updated**: 2026-09-01

---

**Maintained by**: fmen2025  
**Contributors**: Team members involved in server inventory management
