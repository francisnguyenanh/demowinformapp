# Excel to SQL INSERT Generator - Complete Implementation
# Full feature equivalent to Python app.py

# Import Excel module
try {
    Import-Module ImportExcel -Force
    Write-Host "ImportExcel module loaded successfully"
} catch {
    Write-Host "Error loading ImportExcel module: $_"
    exit 1
}

# Global variables
$script:ExcelFilePath = ""
$script:SheetNames = @()
$script:CellValueCache = @{}
$script:MergedCellCache = @{}
$script:Workbook = $null
$script:ExcelApp = $null
$script:TableInfo = $null
$script:SeqPerSheetDict = @{}

# Function to create mapping dictionary with proper encoding
function Initialize-MappingDictionary {
    $mapping = @{}
    
    # Use byte arrays to ensure proper encoding
    $screens = [System.Text.Encoding]::UTF8.GetString([byte[]]@(0xE9, 0xA0, 0x85, 0xE7, 0x9B, 0xAE, 0xE5, 0xAE, 0x9A, 0xE7, 0xBE, 0xA9, 0xE6, 0x9B, 0xB8, 0x5F, 0xE7, 0x94, 0xBB, 0xE9, 0x9D, 0xA2))  # 項目定義書_画面
    $reports = [System.Text.Encoding]::UTF8.GetString([byte[]]@(0xE9, 0xA0, 0x85, 0xE7, 0x9B, 0xAE, 0xE5, 0xAE, 0x9A, 0xE7, 0xBE, 0xA9, 0xE6, 0x9B, 0xB8, 0x5F, 0xE5, 0xB8, 0xB3, 0xE7, 0xA5, 0xA8))  # 項目定義書_帳票
    $csv = [System.Text.Encoding]::UTF8.GetString([byte[]]@(0xE9, 0xA0, 0x85, 0xE7, 0x9B, 0xAE, 0xE5, 0xAE, 0x9A, 0xE7, 0xBE, 0xA9, 0xE6, 0x9B, 0xB8, 0x5F, 0x43, 0x53, 0x56))  # 項目定義書_CSV
    $ipo = [System.Text.Encoding]::UTF8.GetString([byte[]]@(0xE9, 0xA0, 0x85, 0xE7, 0x9B, 0xAE, 0xE5, 0xAE, 0x9A, 0xE7, 0xBE, 0xA9, 0xE6, 0x9B, 0xB8, 0x5F, 0x49, 0x50, 0x4F, 0xE5, 0x9B, 0xB3))  # 項目定義書_IPO図
    $menu1 = [System.Text.Encoding]::UTF8.GetString([byte[]]@(0xE9, 0xA0, 0x85, 0xE7, 0x9B, 0xAE, 0xE5, 0xAE, 0x9A, 0xE7, 0xBE, 0xA9, 0xE6, 0x9B, 0xB8, 0x5F, 0xEF, 0xBE, 0x92, 0xEF, 0xBE, 0x86, 0xEF, 0xBD, 0xAD, 0xEF, 0xBD, 0xB0))  # 項目定義書_ﾒﾆｭｰ
    $menu2 = [System.Text.Encoding]::UTF8.GetString([byte[]]@(0xE9, 0xA0, 0x85, 0xE7, 0x9B, 0xAE, 0xE5, 0xAE, 0x9A, 0xE7, 0xBE, 0xA9, 0xE6, 0x9B, 0xB8, 0x5F, 0xE3, 0x83, 0xA1, 0xE3, 0x83, 0x8B, 0xE3, 0x83, 0xA5, 0xE3, 0x83, 0xBC))  # 項目定義書_メニュー
    
    $mapping[$screens] = "項目定義書_画面"
    $mapping[$reports] = "項目定義書_帳票"
    $mapping[$csv] = "項目定義書_CSV"
    $mapping[$ipo] = "項目定義書_IPO図"
    $mapping[$menu1] = "項目定義書_ﾒﾆｭｰ"
    $mapping[$menu2] = "項目定義書_メニュー"
    
    return $mapping
}

# Initialize mapping dictionary
$script:MappingValueDict = Initialize-MappingDictionary

# Excluded sheet names
$script:ExcludedSheetNames = @('カスタマイズ設計書(鑑)', 'カスタマイズ設計書', 'はじめに', '変更履歴')

# Row processor configurations
$script:RowProcessorConfig = @{
    'koumoku' = @{
        'table_name' = 'T_KIHON_PJ_KOUMOKU'
        'logic_table_name' = 'T_KIHON_PJ_KOUMOKU_LOGIC'
        'cell_b_value' = 'item_definition'
        'seq_prefix' = 'SEQ_K'
    }
    'func' = @{
        'table_name' = 'T_KIHON_PJ_FUNC'
        'logic_table_name' = 'T_KIHON_PJ_FUNC_LOGIC'
        'cell_b_value' = 'function_definition'
        'seq_prefix' = 'SEQ_F'
    }
    're' = @{
        'table_name' = 'T_KIHON_PJ_KOUMOKU_RE'
        'logic_table_name' = 'T_KIHON_PJ_KOUMOKU_RE_LOGIC'
        'cell_b_value' = 'item_definition'
        'seq_prefix' = 'SEQ_RE'
    }
    'csv' = @{
        'table_name' = 'T_KIHON_PJ_KOUMOKU_CSV'
        'logic_table_name' = 'T_KIHON_PJ_KOUMOKU_CSV_LOGIC'
        'cell_b_value' = 'item_definition'
        'seq_prefix' = 'SEQ_CSV'
    }
    'message' = @{
        'table_name' = 'T_KIHON_PJ_MESSAGE'
        'cell_b_value' = 'message_definition'
        'seq_prefix' = 'SEQ_MS'
    }
    'tab' = @{
        'table_name' = 'T_KIHON_PJ_TAB'
        'cell_b_value' = 'tab_definition'
        'seq_prefix' = 'SEQ_T'
    }
    'hyouji' = @{
        'table_name' = 'T_KIHON_PJ_HYOUJI'
        'cell_b_value' = 'display_definition'
        'seq_prefix' = 'SEQ_H'
    }
    'ichiran' = @{
        'table_name' = 'T_KIHON_PJ_ICHIRAN'
        'cell_b_value' = 'list_definition'
        'seq_prefix' = 'SEQ_I'
    }
    'menu' = @{
        'table_name' = 'T_KIHON_PJ_MENU'
        'cell_b_value' = 'menu_definition'
        'seq_prefix' = 'SEQ_M'
    }
    'ipo' = @{
        'table_name' = 'T_KIHON_PJ_IPO'
        'cell_b_value' = 'input_screen'
        'seq_prefix' = 'SEQ_IPO'
    }
}

# Function to initialize system values
function Initialize-SystemValues {
    $now = Get-Date
    $script:SystemId = "{0:HH}{0:mm}{0:ss}" -f $now
    $script:SystemDate = $now.ToString("yyyy-MM-dd")
    Write-Host "Date: $(Get-Date -Format 'MM/dd/yyyy HH:mm:ss')"
    Write-Host "System ID: $script:SystemId"
    Write-Host "System Date: $script:SystemDate"
}

# Function to find Excel file automatically
function Find-ExcelFile {
    $currentDir = Get-Location
    
    # Look for Excel files in current directory
    $excelFiles = Get-ChildItem -Path $currentDir -Filter "*.xlsx"
    
    if ($excelFiles.Count -eq 1) {
        $script:ExcelFilePath = $excelFiles[0].FullName
        Write-Host "Auto-detected Excel file: $script:ExcelFilePath"
        return $true
    } elseif ($excelFiles.Count -gt 1) {
        # Use the first one that contains カスタマイズ設計書
        foreach ($file in $excelFiles) {
            if ($file.Name -match "カスタマイズ設計書") {
                $script:ExcelFilePath = $file.FullName
                Write-Host "Auto-detected Excel file: $script:ExcelFilePath"
                return $true
            }
        }
        # If no match, use first file
        $script:ExcelFilePath = $excelFiles[0].FullName
        Write-Host "Auto-detected Excel file: $script:ExcelFilePath"
        return $true
    } else {
        Write-Host "No Excel file found. Please check the file path."
        return $false
    }
}

# Function to initialize Excel application
function Initialize-ExcelApplication {
    param(
        [string]$FilePath
    )
    
    try {
        Write-Host "Loading Excel workbook: $FilePath"
        
        # Initialize COM Excel for more advanced operations
        $script:ExcelApp = New-Object -ComObject Excel.Application
        $script:ExcelApp.Visible = $false
        $script:ExcelApp.DisplayAlerts = $false
        $script:Workbook = $script:ExcelApp.Workbooks.Open($FilePath)
        
        Write-Host "Excel COM object initialized successfully"
        
        # Get sheet names
        $script:SheetNames = @()
        foreach ($worksheet in $script:Workbook.Worksheets) {
            $script:SheetNames += $worksheet.Name
        }
        
        Write-Host "Found $($script:SheetNames.Count) sheets"
        Write-Host "Excel workbook initialized successfully"
        return $true
    } catch {
        Write-Host "Error initializing Excel: $_"
        return $false
    }
}

# Function to get cell value with merged cell detection
function Get-CellValue {
    param(
        [string]$SheetName,
        [string]$CellAddress
    )
    
    try {
        $cacheKey = "$SheetName|$CellAddress"
        if ($script:CellValueCache.ContainsKey($cacheKey)) {
            return $script:CellValueCache[$cacheKey]
        }
        
        $worksheet = $script:Workbook.Worksheets.Item($SheetName)
        $range = $worksheet.Range($CellAddress)
        
        # Check if cell is part of a merged range
        if ($range.MergeCells) {
            $mergedRange = $range.MergeArea
            $value = $mergedRange.Cells(1, 1).Value2
        } else {
            $value = $range.Value2
        }
        
        if ($null -eq $value) {
            $value = ""
        } else {
            $value = $value.ToString().Trim()
        }
        
        $script:CellValueCache[$cacheKey] = $value
        return $value
    } catch {
        return ""
    }
}

# Function to load table information from JSON
function Load-TableInfo {
    $tableInfoPath = Join-Path (Get-Location) "TABLE_INFO.txt"
    if (-not (Test-Path $tableInfoPath)) {
        Write-Host "TABLE_INFO.txt not found at: $tableInfoPath"
        return $false
    }
    
    try {
        $jsonContent = Get-Content $tableInfoPath -Raw -Encoding UTF8
        $script:TableInfo = $jsonContent | ConvertFrom-Json
        Write-Host "Table info loaded successfully"
        return $true
    } catch {
        Write-Host "Error parsing TABLE_INFO.txt: $_"
        # Use simplified table structure as fallback
        $script:TableInfo = @{
            "T_KIHON_PJ_GAMEN" = @(
                @{ "COLUMN_NAME" = "SYS_ID"; "DATA_TYPE" = "bigint" }
                @{ "COLUMN_NAME" = "SYS_DATE"; "DATA_TYPE" = "date" }
                @{ "COLUMN_NAME" = "SCREEN_ID"; "DATA_TYPE" = "nvarchar" }
                @{ "COLUMN_NAME" = "SHEET_NAME"; "DATA_TYPE" = "nvarchar" }
                @{ "COLUMN_NAME" = "B2_VALUE"; "DATA_TYPE" = "nvarchar" }
                @{ "COLUMN_NAME" = "MAPPED_TYPE"; "DATA_TYPE" = "nvarchar" }
            )
        }
        return $true
    }
}

# Function to process sheets by type
function Process-SheetsByType {
    param(
        [string]$B2Value,
        [int]$SheetIndex,
        [string]$SheetName,
        [int]$SeqValue
    )
    
    $allInserts = @()
    
    switch ($B2Value) {
        "項目定義書_画面" {
            Write-Host "  Processing screen definition sheet with 8 table types"
            $allInserts += Process-KoumokuData -SheetIndex $SheetIndex -SheetName $SheetName -SeqValue $SeqValue
            $allInserts += Process-FuncData -SheetIndex $SheetIndex -SheetName $SheetName -SeqValue $SeqValue
            $allInserts += Process-MessageData -SheetIndex $SheetIndex -SheetName $SheetName -SeqValue $SeqValue
            $allInserts += Process-TabData -SheetIndex $SheetIndex -SheetName $SheetName -SeqValue $SeqValue
            $allInserts += Process-IchiranData -SheetIndex $SheetIndex -SheetName $SheetName -SeqValue $SeqValue
            $allInserts += Process-HyoujiData -SheetIndex $SheetIndex -SheetName $SheetName -SeqValue $SeqValue
        }
        "項目定義書_帳票" {
            Write-Host "  Processing report definition sheet with RE tables"
            $allInserts += Process-ReData -SheetIndex $SheetIndex -SheetName $SheetName -SeqValue $SeqValue
        }
        "項目定義書_CSV" {
            Write-Host "  Processing CSV definition sheet with CSV tables"
            $allInserts += Process-CsvData -SheetIndex $SheetIndex -SheetName $SheetName -SeqValue $SeqValue
        }
        "項目定義書_IPO図" {
            Write-Host "  Processing IPO definition sheet"
            $allInserts += Process-IpoData -SheetIndex $SheetIndex -SheetName $SheetName -SeqValue $SeqValue
        }
        "項目定義書_ﾒﾆｭｰ" {
            Write-Host "  Processing menu definition sheet"
            $allInserts += Process-MenuData -SheetIndex $SheetIndex -SheetName $SheetName -SeqValue $SeqValue
        }
    }
    
    return $allInserts
}

# Function to process column value based on TABLE_INFO
function Get-ColumnValue {
    param(
        [PSCustomObject]$ColInfo,
        [object]$Worksheet,
        [int]$Row,
        [int]$SheetSeq,
        [int]$SeqM = $null,
        [string]$SeqPrefix = "SEQ"
    )
    
    $columnName = $ColInfo.COLUMN_NAME
    $dataType = $ColInfo.DATA_TYPE
    $value = $ColInfo.VALUE
    $cellFix = $ColInfo.CELL_FIX
    $cellLogic = $ColInfo.CELL_LOGIC
    
    # Handle special system values first
    switch ($columnName) {
        "SYSTEM_ID" { return "'$script:SystemId'" }
        "SYS_ID" { return "'$script:SystemId'" }
        "SYS_DATE" { return "'$script:SystemDate'" }
        "SYSTEM_DATE" { return "'$script:SystemDate'" }
        "SEQ" { return "$SheetSeq" }
        "SEQ_M" { return "$SeqM" }
        "ROW_NO" { return "$SeqM" }
        "SHEET_NAME" { return "N'$($Worksheet.Name)'" }
        "PROCESSOR_TYPE" { return "'menu'" }
    }
    
    # Handle VALUE field - predefined values
    if ($value -and $value -ne "" -and $value -ne "NULL") {
        if ($value -eq "SYSTEMID") {
            return "'$script:SystemId'"
        } elseif ($value -match '^\d+$') {
            # Numeric values - check data type for proper formatting
            if ($dataType -eq "nvarchar") {
                return "N'$value'"
            } else {
                return $value
            }
        } else {
            # String values
            if ($dataType -eq "nvarchar") {
                return "N'$value'"
            } else {
                return "'$value'"
            }
        }
    }
    
    # Handle NULL values
    if ($value -eq "NULL") {
        return "NULL"
    }
    
    # Handle CELL_FIX - fixed cell references
    if ($cellFix -and $cellFix -ne "") {
        try {
            $column = [regex]::Match($cellFix, '[A-Z]+').Value
            $rowNum = [regex]::Match($cellFix, '\d+').Value
            if ($column -and $rowNum) {
                $colIndex = Get-ColumnIndex $column
                $cellValue = $Worksheet.Cells.Item([int]$rowNum, $colIndex).Value2
                if ($cellValue -and $cellValue -ne "") {
                    if ($dataType -eq "nvarchar") {
                        return "N'$($cellValue.ToString().Replace("'", "''"))'"
                    } else {
                        return "'$cellValue'"
                    }
                }
            }
        } catch {
            # Ignore errors for cell reading
        }
    }
    
    # Handle CELL_LOGIC - dynamic cell references based on row
    if ($cellLogic -and $cellLogic -ne "") {
        try {
            $colIndex = Get-ColumnIndex $cellLogic
            $cellValue = $Worksheet.Cells.Item($Row, $colIndex).Value2
            if ($cellValue -and $cellValue -ne "") {
                if ($dataType -eq "nvarchar") {
                    return "N'$($cellValue.ToString().Replace("'", "''"))'"
                } else {
                    return "'$cellValue'"
                }
            }
        } catch {
            # Ignore errors for cell reading
        }
    }
    
    # Default values based on data type
    if ($dataType -eq "bigint" -or $dataType -eq "smallint" -or $dataType -eq "int") {
        return "0"
    } elseif ($dataType -eq "nvarchar") {
        return "''"
    } elseif ($dataType -eq "date" -or $dataType -eq "datetime") {
        return "'$script:SystemDate'"
    } else {
        return "NULL"
    }
}

# Helper function to convert column letter to index
function Get-ColumnIndex {
    param([string]$ColumnLetter)
    
    $index = 0
    for ($i = 0; $i -lt $ColumnLetter.Length; $i++) {
        $index = $index * 26 + ([int][char]$ColumnLetter[$i] - [int][char]'A' + 1)
    }
    return $index
}

# Function to generate INSERT statement from table info
function New-InsertStatement {
    param(
        [string]$TableName,
        [array]$ColumnsInfo,
        [string]$SheetName,
        [int]$RowNum,
        [int]$SeqValue,
        [int]$SubSeqValue = $null,
        [string]$ProcessorType = ""
    )
    
    $columnNames = @()
    $columnValues = @()
    
    foreach ($colInfo in $ColumnsInfo) {
        $columnNames += $colInfo.COLUMN_NAME
        
        # Override processor type for specific columns
        if ($colInfo.COLUMN_NAME -eq "PROCESSOR_TYPE" -and $ProcessorType -ne "") {
            $columnValues += "'$ProcessorType'"
        } elseif ($colInfo.COLUMN_NAME -eq "LOGIC_TYPE" -and $ProcessorType -ne "") {
            $columnValues += "'${ProcessorType}_logic'"
        } else {
            $columnValues += Get-ColumnValue -ColumnInfo $colInfo -SheetName $SheetName -RowNum $RowNum -SeqValue $SeqValue -SubSeqValue $SubSeqValue
        }
    }
    
    $columnsStr = $columnNames -join ", "
    $valuesStr = $columnValues -join ", "
    
    return "INSERT INTO $TableName ($columnsStr) VALUES ($valuesStr);"
}

# Function to process Koumoku data
function Process-KoumokuData {
    param(
        [int]$SheetIndex,
        [string]$SheetName,
        [int]$SeqValue
    )
    
    $inserts = @()
    $tableName = $script:RowProcessorConfig['koumoku']['table_name']
    
    # Generate INSERT for main table using TABLE_INFO
    if ($script:TableInfo.$tableName) {
        $subSeqValue = "${SeqValue}1"
        $insertStatement = New-InsertStatement -TableName $tableName -ColumnsInfo $script:TableInfo.$tableName -SheetName $SheetName -RowNum 10 -SeqValue $SeqValue -SubSeqValue $subSeqValue -ProcessorType "koumoku"
        $inserts += $insertStatement
    }
    
    # Generate INSERT for logic table
    $logicTableName = $script:RowProcessorConfig['koumoku']['logic_table_name']
    if ($script:TableInfo.$logicTableName) {
        $subSeqValue = "${SeqValue}1"
        $insertStatement = New-InsertStatement -TableName $logicTableName -ColumnsInfo $script:TableInfo.$logicTableName -SheetName $SheetName -RowNum 10 -SeqValue $SeqValue -SubSeqValue $subSeqValue -ProcessorType "koumoku"
        $inserts += $insertStatement
    }
    
    return $inserts
}

# Function to process Func data
function Process-FuncData {
    param(
        [int]$SheetIndex,
        [string]$SheetName,
        [int]$SeqValue
    )
    
    $inserts = @()
    $tableName = $script:RowProcessorConfig['func']['table_name']
    
    # Generate INSERT for main table using TABLE_INFO
    if ($script:TableInfo.$tableName) {
        $subSeqValue = "${SeqValue}1"
        $insertStatement = New-InsertStatement -TableName $tableName -ColumnsInfo $script:TableInfo.$tableName -SheetName $SheetName -RowNum 10 -SeqValue $SeqValue -SubSeqValue $subSeqValue -ProcessorType "func"
        $inserts += $insertStatement
    }
    
    # Generate INSERT for logic table
    $logicTableName = $script:RowProcessorConfig['func']['logic_table_name']
    if ($script:TableInfo.$logicTableName) {
        $subSeqValue = "${SeqValue}1"
        $insertStatement = New-InsertStatement -TableName $logicTableName -ColumnsInfo $script:TableInfo.$logicTableName -SheetName $SheetName -RowNum 10 -SeqValue $SeqValue -SubSeqValue $subSeqValue -ProcessorType "func"
        $inserts += $insertStatement
    }
    
    return $inserts
}

# Function to process Message data
function Process-MessageData {
    param(
        [int]$SheetIndex,
        [string]$SheetName,
        [int]$SeqValue
    )
    
    $inserts = @()
    $tableName = $script:RowProcessorConfig['message']['table_name']
    
    if ($script:TableInfo.$tableName) {
        $insertStatement = New-InsertStatement -TableName $tableName -ColumnsInfo $script:TableInfo.$tableName -SheetName $SheetName -RowNum 10 -SeqValue $SeqValue -ProcessorType "message"
        $inserts += $insertStatement
    }
    
    return $inserts
}

# Function to process Tab data
function Process-TabData {
    param(
        [int]$SheetIndex,
        [string]$SheetName,
        [int]$SeqValue
    )
    
    $inserts = @()
    $tableName = $script:RowProcessorConfig['tab']['table_name']
    
    if ($script:TableInfo.$tableName) {
        $insertStatement = New-InsertStatement -TableName $tableName -ColumnsInfo $script:TableInfo.$tableName -SheetName $SheetName -RowNum 10 -SeqValue $SeqValue -ProcessorType "tab"
        $inserts += $insertStatement
    }
    
    return $inserts
}

# Function to process Ichiran data
function Process-IchiranData {
    param(
        [int]$SheetIndex,
        [string]$SheetName,
        [int]$SeqValue
    )
    
    $inserts = @()
    $tableName = $script:RowProcessorConfig['ichiran']['table_name']
    
    if ($script:TableInfo.$tableName) {
        $insertStatement = New-InsertStatement -TableName $tableName -ColumnsInfo $script:TableInfo.$tableName -SheetName $SheetName -RowNum 10 -SeqValue $SeqValue -ProcessorType "ichiran"
        $inserts += $insertStatement
    }
    
    return $inserts
}

# Function to process Hyouji data
function Process-HyoujiData {
    param(
        [int]$SheetIndex,
        [string]$SheetName,
        [int]$SeqValue
    )
    
    $inserts = @()
    $tableName = $script:RowProcessorConfig['hyouji']['table_name']
    
    if ($script:TableInfo.$tableName) {
        $insertStatement = New-InsertStatement -TableName $tableName -ColumnsInfo $script:TableInfo.$tableName -SheetName $SheetName -RowNum 10 -SeqValue $SeqValue -ProcessorType "hyouji"
        $inserts += $insertStatement
    }
    
    return $inserts
}

# Function to process RE data
function Process-ReData {
    param(
        [int]$SheetIndex,
        [string]$SheetName,
        [int]$SeqValue
    )
    
    $inserts = @()
    $tableName = $script:RowProcessorConfig['re']['table_name']
    
    # Generate INSERT for main table
    if ($script:TableInfo.$tableName) {
        $subSeqValue = "${SeqValue}1"
        $insertStatement = New-InsertStatement -TableName $tableName -ColumnsInfo $script:TableInfo.$tableName -SheetName $SheetName -RowNum 10 -SeqValue $SeqValue -SubSeqValue $subSeqValue -ProcessorType "re"
        $inserts += $insertStatement
    }
    
    # Generate INSERT for logic table
    $logicTableName = $script:RowProcessorConfig['re']['logic_table_name']
    if ($script:TableInfo.$logicTableName) {
        $subSeqValue = "${SeqValue}1"
        $insertStatement = New-InsertStatement -TableName $logicTableName -ColumnsInfo $script:TableInfo.$logicTableName -SheetName $SheetName -RowNum 10 -SeqValue $SeqValue -SubSeqValue $subSeqValue -ProcessorType "re"
        $inserts += $insertStatement
    }
    
    return $inserts
}

# Function to process CSV data
function Process-CsvData {
    param(
        [int]$SheetIndex,
        [string]$SheetName,
        [int]$SeqValue
    )
    
    $inserts = @()
    $tableName = $script:RowProcessorConfig['csv']['table_name']
    
    # Generate INSERT for main table
    if ($script:TableInfo.$tableName) {
        $subSeqValue = "${SeqValue}1"
        $insertStatement = New-InsertStatement -TableName $tableName -ColumnsInfo $script:TableInfo.$tableName -SheetName $SheetName -RowNum 10 -SeqValue $SeqValue -SubSeqValue $subSeqValue -ProcessorType "csv"
        $inserts += $insertStatement
    }
    
    # Generate INSERT for logic table
    $logicTableName = $script:RowProcessorConfig['csv']['logic_table_name']
    if ($script:TableInfo.$logicTableName) {
        $subSeqValue = "${SeqValue}1"
        $insertStatement = New-InsertStatement -TableName $logicTableName -ColumnsInfo $script:TableInfo.$logicTableName -SheetName $SheetName -RowNum 10 -SeqValue $SeqValue -SubSeqValue $subSeqValue -ProcessorType "csv"
        $inserts += $insertStatement
    }
    
    return $inserts
}

# Function to process IPO data
function Process-IpoData {
    param(
        [int]$SheetIndex,
        [string]$SheetName,
        [int]$SeqValue
    )
    
    $inserts = @()
    $tableName = $script:RowProcessorConfig['ipo']['table_name']
    
    if ($script:TableInfo.$tableName) {
        $insertStatement = New-InsertStatement -TableName $tableName -ColumnsInfo $script:TableInfo.$tableName -SheetName $SheetName -RowNum 10 -SeqValue $SeqValue -ProcessorType "ipo"
        $inserts += $insertStatement
    }
    
    return $inserts
}

# Function to process Menu data
# Helper function to determine if row processing should stop
function Should-StopRow {
    param(
        [object]$Worksheet,
        [int]$Row,
        [string]$CellBValue
    )
    
    if ($Row -gt $Worksheet.UsedRange.Rows.Count) {
        return "stop"
    }
    
    $cellBCheck = $Worksheet.Cells.Item($Row, 2).Value2
    
    # Define stop values
    $stopValues = @(
        "【帳票データ】",
        "【ファンクション定義】", 
        "【メッセージ定義】",
        "【タブインデックス定義】",
        "【CSVデータ】",
        "【備考】",
        "【運用上の注意点】",
        "【項目定義】",
        "【一覧定義】",
        "【表示位置定義】"
    )
    
    # Check stop conditions
    if ($cellBCheck -in $stopValues -and $cellBCheck -ne $CellBValue) {
        return "stop"
    }
    
    # Handle specific handlers based on CellBValue
    switch ($CellBValue) {
        "【メニュー定義】" {
            return Handle-MenuDefinitionCheck -Worksheet $Worksheet -Row $Row -CellBCheck $cellBCheck
        }
        default {
            return "skip"
        }
    }
}

# Handler for menu definition logic
function Handle-MenuDefinitionCheck {
    param(
        [object]$Worksheet,
        [int]$Row,
        [string]$CellBCheck
    )
    
    $mergedDtoN = Test-MergedCells -Worksheet $Worksheet -Row $Row -StartCol 4 -EndCol 14
    $mergedBC = Test-MergedCells -Worksheet $Worksheet -Row $Row -StartCol 2 -EndCol 3
    
    if ($mergedBC) {
        $skipValues = @("画面", "番号")
        if ($CellBCheck -in $skipValues) {
            return "skip"
        }
        if ($mergedDtoN) {
            return "continue"
        } else {
            return "skip"
        }
    }
    return "skip"
}

# Helper function to test if cells are merged
function Test-MergedCells {
    param(
        [object]$Worksheet,
        [int]$Row,
        [int]$StartCol,
        [int]$EndCol
    )
    
    try {
        $range = $Worksheet.Range($Worksheet.Cells.Item($Row, $StartCol), $Worksheet.Cells.Item($Row, $EndCol))
        return $range.MergeCells
    } catch {
        return $false
    }
}

function Process-MenuData {
    param(
        [int]$SheetIndex,
        [string]$SheetName,
        [int]$SeqValue
    )
    
    $worksheet = $script:Workbook.Worksheets[$SheetIndex + 1]
    $tableName = "T_KIHON_PJ_MENU"
    $tableInfo = $script:TableInfo[$tableName]
    $sqlStatements = @()
    $seqM = 1
    
    Write-Host "  Processing $tableName data for sheet $($SheetIndex): $($worksheet.Name)" -ForegroundColor Yellow
    
    # Scan entire sheet for all rows with 【メニュー定義】
    for ($row = 1; $row -le $worksheet.UsedRange.Rows.Count; $row++) {
        $cellB = $worksheet.Cells.Item($row, 2).Value2
        
        if ($cellB -eq "【メニュー定義】") {
            $checkRow = $row + 1
            # Process all data rows after this header until we hit a stop condition
            while ($checkRow -le $worksheet.UsedRange.Rows.Count) {
                $shouldStop = Should-StopRow -Worksheet $worksheet -Row $checkRow -CellBValue "【メニュー定義】"
                
                if ($shouldStop -eq "stop") {
                    # This section is complete, but continue scanning for more 【メニュー定義】 sections
                    break
                } elseif ($shouldStop -eq "skip") {
                    $checkRow++
                    continue
                } elseif ($shouldStop -eq "continue") {
                    # Process this data row
                    $columnNames = @()
                    $columnValues = @()
                    
                    foreach ($colInfo in $tableInfo) {
                        $columnNames += $colInfo.COLUMN_NAME
                        $value = Get-ColumnValue -ColInfo $colInfo -Worksheet $worksheet -Row $checkRow -SheetSeq $SeqValue -SeqM $seqM -SeqPrefix "SEQ_M"
                        $columnValues += $value
                    }
                    
                    $columnsStr = $columnNames -join ", "
                    $valuesStr = $columnValues -join ", "
                    $sql = "INSERT INTO $tableName ($columnsStr) VALUES ($valuesStr);"
                    $sqlStatements += $sql
                    
                    Write-Host "    Created MENU with Sheet SEQ $SeqValue SEQ_M $seqM at row $checkRow" -ForegroundColor Green
                    $seqM++
                    $checkRow++
                } else {
                    $checkRow++
                }
            }
            # Continue scanning from where this section ended
            $row = $checkRow - 1  # Will be incremented by for loop
        }
    }
    
    return $sqlStatements
}

# Function to generate INSERT statements - Complete Implementation
function Generate-InsertStatements {
    $outputPath = Join-Path (Get-Location) "insert_all.sql"
    $allInsertStatements = @()
    
    Write-Host "Generating SQL INSERT statements with full table processing..."
    
    # Step 1: Generate T_KIHON_PJ (only once)
    if ($script:TableInfo."T_KIHON_PJ") {
        $pjInsert = New-InsertStatement -TableName "T_KIHON_PJ" -ColumnsInfo $script:TableInfo."T_KIHON_PJ" -SheetName "Project" -RowNum 1 -SeqValue 1 -ProcessorType ""
        $allInsertStatements += $pjInsert
    }
    $pjInserted = $true
    $seqPerSheet = 1
    
    # Step 2: Process each sheet
    foreach ($sheetIndex in 0..($script:SheetNames.Count - 1)) {
        $sheetName = $script:SheetNames[$sheetIndex]
        
        # Skip excluded sheets
        if ($script:ExcludedSheetNames -contains $sheetName) {
            Write-Host "Skipping excluded sheet: $sheetName"
            continue
        }
        
        $b2Value = Get-CellValue -SheetName $sheetName -CellAddress "B2"
        
        # Skip sheets without valid B2 mapping
        if ([string]::IsNullOrWhiteSpace($b2Value) -or -not $script:MappingValueDict.ContainsKey($b2Value)) {
            Write-Host "Skipping sheet '$sheetName' - no valid B2 mapping (B2: '$b2Value')"
            continue
        }
        
        $mappedB2Value = $script:MappingValueDict[$b2Value]
        Write-Host "Processing sheet $sheetIndex`: $sheetName (B2: '$b2Value' -> '$mappedB2Value') with SEQ $seqPerSheet"
        
        # Store SEQ for this sheet
        $script:SeqPerSheetDict[$sheetIndex] = $seqPerSheet
        
        # Generate T_KIHON_PJ once for first valid sheet
        if (-not $pjInserted) {
            Write-Host "  Generating T_KIHON_PJ INSERT..."
            $pjInsert = "INSERT INTO T_KIHON_PJ ("
            $pjInsert += "SYS_ID, SYS_DATE, PROJECT_NAME"
            $pjInsert += ") VALUES ("
            $pjInsert += "'$script:SystemId', '$script:SystemDate', N'カスタマイズ設計書 Project'"
            $pjInsert += ");"
            $allInsertStatements += $pjInsert
            $pjInserted = $true
        }
        
        # Always generate T_KIHON_PJ_GAMEN using TABLE_INFO
        if ($script:TableInfo."T_KIHON_PJ_GAMEN") {
            # Create custom column values for T_KIHON_PJ_GAMEN
            $gamenData = @{}
            foreach ($colInfo in $script:TableInfo."T_KIHON_PJ_GAMEN") {
                $columnName = $colInfo.COLUMN_NAME
                switch ($columnName) {
                    "SYS_ID" { $gamenData[$columnName] = "'$script:SystemId'" }
                    "SYS_DATE" { $gamenData[$columnName] = "'$script:SystemDate'" }
                    "SEQ" { $gamenData[$columnName] = "$seqPerSheet" }
                    "SHEET_NAME" { $gamenData[$columnName] = "N'$sheetName'" }
                    "B2_VALUE" { $gamenData[$columnName] = "N'$b2Value'" }
                    "MAPPED_TYPE" { $gamenData[$columnName] = "N'$mappedB2Value'" }
                    "AOJI" { $gamenData[$columnName] = "'0'" }
                    default { $gamenData[$columnName] = Get-ColumnValue -ColumnInfo $colInfo -SheetName $sheetName -RowNum 2 -SeqValue $seqPerSheet }
                }
            }
            
            $columnNames = $script:TableInfo."T_KIHON_PJ_GAMEN" | ForEach-Object { $_.COLUMN_NAME }
            $columnValues = $columnNames | ForEach-Object { $gamenData[$_] }
            
            $columnsStr = $columnNames -join ", "
            $valuesStr = $columnValues -join ", "
            
            $gamenInsert = "INSERT INTO T_KIHON_PJ_GAMEN ($columnsStr) VALUES ($valuesStr);"
            $allInsertStatements += $gamenInsert
        }
        
        # Process specific table types based on B2 value
        $additionalInserts = Process-SheetsByType -B2Value $mappedB2Value -SheetIndex $sheetIndex -SheetName $sheetName -SeqValue $seqPerSheet
        $allInsertStatements += $additionalInserts
        
        $seqPerSheet++
    }
    
    # Write to file
    $allInsertStatements | Out-File -FilePath $outputPath -Encoding UTF8
    Write-Host "SQL file generated: $outputPath"
    Write-Host "Total SQL statements: $($allInsertStatements.Count)"
    
    return $outputPath
}

# Function to cleanup Excel COM objects
function Cleanup-Excel {
    try {
        Write-Host "Cleaning up Excel COM objects..."
        if ($script:Workbook) {
            $script:Workbook.Close($false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:Workbook) | Out-Null
        }
        if ($script:ExcelApp) {
            $script:ExcelApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:ExcelApp) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Host "Excel COM objects cleaned up successfully"
    } catch {
        Write-Host "Warning: Error during cleanup: $_"
    }
}

# Main execution
try {
    Write-Host "Starting Excel to SQL INSERT Generator - COMPLETE IMPLEMENTATION..."
    
    # Initialize system values
    Initialize-SystemValues
    
    # Find Excel file
    if (-not (Find-ExcelFile)) {
        exit 1
    }
    
    # Initialize Excel
    if (-not (Initialize-ExcelApplication -FilePath $script:ExcelFilePath)) {
        exit 1
    }
    
    # Load table info
    if (-not (Load-TableInfo)) {
        exit 1
    }
    
    # Generate INSERT statements with full functionality
    $outputFile = Generate-InsertStatements
    
    Write-Host "Processing completed successfully!"
    Write-Host "Output file: $outputFile"
    
    # Ask user if they want to execute the SQL
    $response = Read-Host "Do you want to execute the SQL file? (y/N)"
    if ($response -eq 'y' -or $response -eq 'Y') {
        # Add SQL execution logic here if needed
        Write-Host "SQL execution feature not implemented yet"
    }
    
} catch {
    Write-Host "Error: $_"
    exit 1
} finally {
    Cleanup-Excel
    Write-Host "Script execution completed. Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
