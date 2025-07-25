# =====================================================================
# SQL Script Executor (PowerShell Version)
# Executes all SQL scripts in output_scripts_update directory
# =====================================================================

param(
    [string]$ConnectStringFile = "connect_string.txt",
    [string]$ScriptDirectory = "output_scripts_update",
    [string]$OutputFile = "output.txt"
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

# Function to execute SQL script using sqlcmd
function Execute-SqlScript {
    param(
        [string]$ScriptPath,
        [string]$Server,
        [string]$Database,
        [string]$User,
        [string]$Password
    )
    
    try {
        # Build sqlcmd command
        $sqlcmdArgs = @(
            "-S", $Server,
            "-d", $Database,
            "-U", $User,
            "-P", $Password,
            "-i", $ScriptPath,
            "-f", "65001"
        )
        
        # Execute sqlcmd
        $result = & sqlcmd @sqlcmdArgs 2>&1
        
        return @{
            Success = $LASTEXITCODE -eq 0
            Output = $result -join "`n"
            ExitCode = $LASTEXITCODE
        }
    }
    catch {
        return @{
            Success = $false
            Output = "Error executing sqlcmd: $_"
            ExitCode = -1
        }
    }
}

# Main script
try {
    Write-Host "Starting SQL script execution..." -ForegroundColor Green
    Write-Host "Date: $(Get-Date)" -ForegroundColor Gray
    Write-Host ""
    
    # Initialize output file
    $outputContent = @()
    $outputContent += "Starting SQL script execution..."
    $outputContent += "Date: $(Get-Date)"
    $outputContent += ""
    
    # Read and parse connection string
    Write-Host "Reading connection string from $ConnectStringFile..." -ForegroundColor Cyan
    $odbcConnString = Read-ConnectString -FilePath $ConnectStringFile
    $connParams = Parse-ConnectString -ConnString $odbcConnString
    
    Write-Host "Server: $($connParams['SERVER'])" -ForegroundColor Gray
    Write-Host "Database: $($connParams['DATABASE'])" -ForegroundColor Gray
    Write-Host "User: $($connParams['UID'])" -ForegroundColor Gray
    Write-Host ""
    
    # Check if script directory exists
    if (-not (Test-Path $ScriptDirectory)) {
        $errorMsg = "Error: Script directory '$ScriptDirectory' does not exist."
        Write-Host $errorMsg -ForegroundColor Red
        $outputContent += $errorMsg
        $outputContent | Out-File -FilePath $OutputFile -Encoding UTF8
        Read-Host "Press Enter to continue"
        exit 1
    }
    
    # Get all SQL files in the directory
    $sqlFiles = Get-ChildItem -Path $ScriptDirectory -Filter "*.sql" | Sort-Object Name
    
    if ($sqlFiles.Count -eq 0) {
        $warningMsg = "Warning: No SQL files found in directory '$ScriptDirectory'"
        Write-Host $warningMsg -ForegroundColor Yellow
        $outputContent += $warningMsg
        $outputContent | Out-File -FilePath $OutputFile -Encoding UTF8
        Read-Host "Press Enter to continue"
        exit 0
    }
    
    Write-Host "Found $($sqlFiles.Count) SQL files to execute" -ForegroundColor Green
    Write-Host ""
    
    # Execute each SQL file
    $successCount = 0
    $errorCount = 0
    
    foreach ($sqlFile in $sqlFiles) {
        $fileName = $sqlFile.Name
        Write-Host "Running $fileName..." -ForegroundColor Cyan
        
        $outputContent += "Running $fileName..."
        
        # Execute the SQL script
        $result = Execute-SqlScript -ScriptPath $sqlFile.FullName -Server $connParams['SERVER'] -Database $connParams['DATABASE'] -User $connParams['UID'] -Password $connParams['PWD']
        
        if ($result.Success) {
            Write-Host "  $fileName executed successfully" -ForegroundColor Green
            $outputContent += "$fileName executed successfully"
            $successCount++
        } else {
            Write-Host "  Warning: $fileName had errors - possibly table does not exist in target DB" -ForegroundColor Yellow
            Write-Host "  Continuing with next file..." -ForegroundColor Yellow
            $outputContent += "Warning: $fileName had errors - possibly table does not exist in target DB"
            $outputContent += "Continuing with next file..."
            $errorCount++
        }
        
        # Add command output to log
        if (-not [string]::IsNullOrWhiteSpace($result.Output)) {
            $outputContent += "Output:"
            $outputContent += $result.Output
        }
        
        $outputContent += ""
        Write-Host ""
    }
    
    # Summary
    Write-Host "Execution Summary:" -ForegroundColor Cyan
    Write-Host "  Total files: $($sqlFiles.Count)" -ForegroundColor Gray
    Write-Host "  Successful: $successCount" -ForegroundColor Green
    Write-Host "  With errors: $errorCount" -ForegroundColor Yellow
    Write-Host ""
    
    $outputContent += "Execution Summary:"
    $outputContent += "Total files: $($sqlFiles.Count)"
    $outputContent += "Successful: $successCount"
    $outputContent += "With errors: $errorCount"
    $outputContent += ""
    
    if ($errorCount -eq 0) {
        Write-Host "All scripts executed successfully!" -ForegroundColor Green
        $outputContent += "All scripts executed successfully!"
    } else {
        Write-Host "Script execution completed with some warnings." -ForegroundColor Yellow
        $outputContent += "Script execution completed with some warnings."
    }
    
    $outputContent += "End time: $(Get-Date)"
    
    # Write output to file
    $outputContent | Out-File -FilePath $OutputFile -Encoding UTF8
    
    Write-Host "Check '$OutputFile' for execution details." -ForegroundColor Cyan
    
}
catch {
    $errorMsg = "Error: $_"
    Write-Host $errorMsg -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    
    # Write error to output file
    $errorContent = @()
    $errorContent += "Error occurred during execution:"
    $errorContent += $errorMsg
    $errorContent += "Stack trace:"
    $errorContent += $_.ScriptStackTrace
    $errorContent += "End time: $(Get-Date)"
    
    $errorContent | Out-File -FilePath $OutputFile -Encoding UTF8
}

Write-Host ""
Write-Host "Script execution completed. Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
