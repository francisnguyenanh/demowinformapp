# =====================================================================
# UPDATE Script Generator (PowerShell Version)
# =====================================================================

param(
    [string]$ConnectStringFile = "connect_string.txt"
)

# Function to read connection string from file
function Read-ConnectString {
    param([string]$FilePath)
    
    if (-not (Test-Path $FilePath)) {
        throw "Connection string file not found: $FilePath"
    }
    
    $content = Get-Content $FilePath -Raw -Encoding UTF8
    return $content.Trim()
}

# Function to parse connection string
function Parse-ConnectString {
    param([string]$ConnString)
    
    $params = @{}
    $parts = $ConnString -split ";"
    
    foreach ($part in $parts) {
        if ($part -match "^(.+?)=(.+)$") {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            $params[$key] = $value
        }
    }
    
    return $params
}

# Function to convert ODBC connection string to ADO.NET format
function Convert-ToAdoNetConnectionString {
    param([hashtable]$ConnParams)
    
    $adoNetConnString = ""
    
    # Map ODBC parameters to ADO.NET parameters
    if ($ConnParams.ContainsKey("SERVER")) {
        $adoNetConnString += "Server=$($ConnParams['SERVER']);"
    }
    if ($ConnParams.ContainsKey("DATABASE")) {
        $adoNetConnString += "Database=$($ConnParams['DATABASE']);"
    }
    if ($ConnParams.ContainsKey("UID")) {
        $adoNetConnString += "User Id=$($ConnParams['UID']);"
    }
    if ($ConnParams.ContainsKey("PWD")) {
        $adoNetConnString += "Password=$($ConnParams['PWD']);"
    }
    
    # Add default settings for ADO.NET
    $adoNetConnString += "Trusted_Connection=false;"
    $adoNetConnString += "MultipleActiveResultSets=true;"
    $adoNetConnString += "Connection Timeout=30;"
    
    return $adoNetConnString.TrimEnd(';')
}

# Function to escape SQL string values
function Escape-SqlString {
    param([string]$Value)
    
    if ([string]::IsNullOrEmpty($Value)) {
        return "NULL"
    }
    
    # Escape single quotes and handle special characters
    $escaped = $Value.Replace("'", "''").Replace("`r", "").Replace("`n", " ")
    return "N'$escaped'"
}

# Main script
try {
    Write-Host "Starting UPDATE script generation..." -ForegroundColor Green
    Write-Host "Date: $(Get-Date)" -ForegroundColor Gray
    Write-Host ""
    
    # Create output directory
    $outputDir = "output_scripts_update"
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir | Out-Null
        Write-Host "Created output directory: $outputDir" -ForegroundColor Yellow
    }
    
    # Read and parse connection string
    Write-Host "Reading connection string from $ConnectStringFile..." -ForegroundColor Cyan
    $odbcConnString = Read-ConnectString -FilePath $ConnectStringFile
    $connParams = Parse-ConnectString -ConnString $odbcConnString
    
    # Convert ODBC format to ADO.NET format
    $connString = Convert-ToAdoNetConnectionString -ConnParams $connParams
    
    Write-Host "Server: $($connParams['SERVER'])" -ForegroundColor Gray
    Write-Host "Database: $($connParams['DATABASE'])" -ForegroundColor Gray
    Write-Host "User: $($connParams['UID'])" -ForegroundColor Gray
    Write-Host "ADO.NET Connection String: $connString" -ForegroundColor Gray
    Write-Host ""
    
    # Load System.Data assembly
    Add-Type -AssemblyName System.Data
    
    # Create connection
    Write-Host "Testing database connection..." -ForegroundColor Cyan
    $connection = New-Object System.Data.SqlClient.SqlConnection($connString)
    $connection.Open()
    Write-Host "Database connection successful!" -ForegroundColor Green
    
    # Get list of tables starting with M_
    Write-Host "Getting list of tables starting with 'M_'..." -ForegroundColor Cyan
    
    $tableQuery = @"
SELECT TABLE_NAME 
FROM INFORMATION_SCHEMA.TABLES 
WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME LIKE 'M_%'
ORDER BY TABLE_NAME
"@
    
    $command = $connection.CreateCommand()
    $command.CommandText = $tableQuery
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
    $tableDataSet = New-Object System.Data.DataSet
    $adapter.Fill($tableDataSet) | Out-Null
    
    $tables = $tableDataSet.Tables[0]
    
    if ($tables.Rows.Count -eq 0) {
        Write-Host "No tables starting with 'M_' found in database" -ForegroundColor Yellow
        return
    }
    
    Write-Host "Found $($tables.Rows.Count) tables to process" -ForegroundColor Green
    Write-Host ""
    
    # Process each table
    foreach ($tableRow in $tables.Rows) {
        $tableName = $tableRow["TABLE_NAME"]
        Write-Host "Processing table: $tableName" -ForegroundColor Cyan
        
        # Get target columns
        $columnQuery = @"
SELECT COLUMN_NAME 
FROM INFORMATION_SCHEMA.COLUMNS 
WHERE TABLE_NAME = '$tableName' AND (
    (COLUMN_NAME LIKE '%NAME%' OR
    COLUMN_NAME LIKE '%TEL%' OR
    COLUMN_NAME LIKE '%FAX%' OR
    COLUMN_NAME LIKE '%POST%' OR
    COLUMN_NAME LIKE '%ADDRESS%' OR
    COLUMN_NAME LIKE '%TANTOU%' OR
    COLUMN_NAME LIKE '%CREATE_USER%' OR
    COLUMN_NAME LIKE '%UPDATE_USER%' OR
    COLUMN_NAME LIKE '%FURIGANA%')
    AND COLUMN_NAME NOT LIKE '%CD%'
)
ORDER BY COLUMN_NAME
"@
        
        $command.CommandText = $columnQuery
        $columnAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
        $columnDataSet = New-Object System.Data.DataSet
        $columnAdapter.Fill($columnDataSet) | Out-Null
        
        $targetColumns = $columnDataSet.Tables[0]
        
        if ($targetColumns.Rows.Count -eq 0) {
            Write-Host "  Table $tableName has no target columns containing required keywords" -ForegroundColor Yellow
            continue
        }
        
        # Build column list
        $columnList = @()
        foreach ($colRow in $targetColumns.Rows) {
            $columnList += $colRow["COLUMN_NAME"]
        }
        
        Write-Host "  Found $($columnList.Count) target columns: $($columnList -join ', ')" -ForegroundColor Gray
        
        # Get data from target columns
        $dataQuery = "SELECT $($columnList -join ', ') FROM $tableName"
        
        try {
            $command.CommandText = $dataQuery
            $dataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
            $dataDataSet = New-Object System.Data.DataSet
            $dataAdapter.Fill($dataDataSet) | Out-Null
            
            $dataTable = $dataDataSet.Tables[0]
            
            # Generate UPDATE script
            $scriptFile = Join-Path $outputDir "$tableName`_update.sql"
            
            # Create script content
            $scriptContent = @()
            $scriptContent += "-- UPDATE script for table $tableName (columns containing NAME, TEL, FAX, POST, ADDRESS, TANTOU, CREATE_USER, UPDATE_USER, FURIGANA)"
            $scriptContent += "-- Number of rows: $($dataTable.Rows.Count)"
            $scriptContent += "-- Generated on: $(Get-Date)"
            $scriptContent += ""
            
            # Generate UPDATE statements for each row
            for ($i = 0; $i -lt $dataTable.Rows.Count; $i++) {
                $row = $dataTable.Rows[$i]
                $setClauses = @()
                
                foreach ($column in $columnList) {
                    $value = $row[$column]
                    if ($value -eq [DBNull]::Value -or $null -eq $value) {
                        $setClauses += "$column = NULL"
                    } else {
                        $escapedValue = Escape-SqlString -Value $value.ToString()
                        $setClauses += "$column = $escapedValue"
                    }
                }
                
                $setClause = $setClauses -join ', '
                $rowNumber = $i + 1
                
                $updateSql = @"
;WITH T AS (SELECT *, ROW_NUMBER() OVER (ORDER BY (SELECT 1)) AS rn FROM $tableName)
UPDATE T SET $setClause WHERE rn = $rowNumber;

"@
                $scriptContent += $updateSql
            }
            
            # Write script to file with UTF-8 BOM
            $scriptContent | Out-File -FilePath $scriptFile -Encoding UTF8
            
            Write-Host "  Generated UPDATE script: $scriptFile" -ForegroundColor Green
            
        }
        catch {
            Write-Host "  Error querying data from table $tableName`: $_" -ForegroundColor Red
            continue
        }
    }
    
    Write-Host ""
    Write-Host "All UPDATE scripts generated successfully!" -ForegroundColor Green
    Write-Host "Check '$outputDir' folder for generated scripts." -ForegroundColor Cyan
    
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
}
finally {
    if ($connection -and $connection.State -eq "Open") {
        $connection.Close()
        Write-Host "Database connection closed." -ForegroundColor Gray
    }
}

Write-Host ""
Write-Host "Script execution completed. Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")