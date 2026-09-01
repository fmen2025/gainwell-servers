# Run SQL queries against sts.server and generate summary report

# Read environment variables
$env_file = "c:\projects\GainwellServers_20260820\.env"
$env_vars = @{}

try {
    Get-Content $env_file | Where-Object { $_ -match '^[A-Za-z_]+=' } | ForEach-Object {
        $key, $value = $_ -split '=', 2
        $env_vars[$key.Trim()] = $value.Trim()
    }
    Write-Host "✓ Environment variables loaded"
} catch {
    Write-Host "✗ Error reading .env file: $_"
    exit 1
}

# Build connection string
$server = $env_vars['DB_SERVER']
$database = $env_vars['DB_NAME']
$user = $env_vars['DB_USER']
$password = $env_vars['DB_PASSWORD']

if (-not $server -or -not $database) {
    Write-Host "✗ DB_SERVER or DB_NAME not found in .env"
    exit 1
}

# Create connection string
$connStr = "Server=$server;Database=$database;User Id=$user;Password=$password;Connection Timeout=5;"

Write-Host "Attempting to connect to SQL Server..."
Write-Host "Server: $server"
Write-Host "Database: $database"
Write-Host ""

# Try to execute Query 1 (Count servers found)
try {
    $conn = New-Object System.Data.SqlClient.SqlConnection
    $conn.ConnectionString = $connStr
    $conn.Open()
    Write-Host "✓ Connected to database"
    
    # Read the SQL file and extract queries
    $sql_file = "c:\projects\GainwellServers_20260820\output\sts_server_comparison.sql"
    $sql_content = Get-Content $sql_file -Raw
    
    # Extract Query 1 (from "SELECT COUNT" to first ";")
    $query1_pattern = "SELECT COUNT\(DISTINCT s\.Servername\).*?FROM sts\.server.*?`);$"
    $query1_match = [regex]::Match($sql_content, $query1_pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
    
    if ($query1_match.Success) {
        $query1 = $query1_match.Value
        
        $cmd = New-Object System.Data.SqlClient.SqlCommand
        $cmd.Connection = $conn
        $cmd.CommandText = $query1
        
        $result = $cmd.ExecuteScalar()
        Write-Host "Found in sts.server: $result servers"
    }
    
    $conn.Close()
    
} catch {
    Write-Host "✗ Database connection failed: $_"
    Write-Host ""
    Write-Host "Since database connection failed, generating summary based on available data:"
}

Write-Host ""
Write-Host "========================================================================="
Write-Host "SUMMARY REPORT: Excel ServerNames vs sts.server"
Write-Host "========================================================================="
Write-Host ""
Write-Host "Source: Excel file - Gainwell Servers_20260820.xlsx (Name column)"
Write-Host "Target: SQL Server - sts.server table"
Write-Host "Date: 2026-08-26"
Write-Host ""
Write-Host "Total distinct ServerNames from Excel: 952"
Write-Host ""
Write-Host "SQL Query Files Available:"
Write-Host "  [1] sts_server_comparison.sql - Contains 3 queries"
Write-Host "      - Query 1: Count servers from Excel found in sts.server"
Write-Host "      - Query 2: List servers found in both"
Write-Host "      - Query 3: List servers from Excel NOT in sts.server"
Write-Host ""
Write-Host "Next Steps:"
Write-Host "  1. Open SQL Server Management Studio"
Write-Host "  2. Connect to the database specified in .env"
Write-Host "  3. Open and run: c:\projects\GainwellServers_20260820\output\sts_server_comparison.sql"
Write-Host "  4. Review results to identify:"
Write-Host "     - How many Excel servers exist in sts.server"
Write-Host "     - Which servers are missing from sts.server"
Write-Host ""
Write-Host "========================================================================="
