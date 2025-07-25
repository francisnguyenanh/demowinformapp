@echo off
setlocal enabledelayedexpansion

REM Read connection string from connect_string.txt
set "CONNECT_STRING_FILE=%~dp0connect_string.txt"
if not exist "%CONNECT_STRING_FILE%" (
    echo Error: connect_string.txt not found
    pause
    exit /b 1
)

set /p CONN_STR=<"%CONNECT_STRING_FILE%"

REM Create output directory
set "OUTPUT_DIR=%~dp0output_scripts_update"
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

REM Create temporary SQL files
set "TEMP_TABLES=%TEMP%\tables_list.txt"
set "TEMP_COLUMNS=%TEMP%\columns_list.txt"
set "TEMP_DATA=%TEMP%\data_export.txt"

REM Get list of tables starting with M_
sqlcmd -S "%CONN_STR%" -Q "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME LIKE 'M_%%'" -h -1 -W -s "," > "%TEMP_TABLES%"

REM Process each table
for /f "tokens=*" %%T in (%TEMP_TABLES%) do (
    set "TABLE_NAME=%%T"
    if not "!TABLE_NAME!"=="" (
        echo Processing table: !TABLE_NAME!
        
        REM Get target columns
        sqlcmd -S "%CONN_STR%" -Q "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '!TABLE_NAME!' AND ((COLUMN_NAME LIKE '%%NAME%%' OR COLUMN_NAME LIKE '%%TEL%%' OR COLUMN_NAME LIKE '%%FAX%%' OR COLUMN_NAME LIKE '%%POST%%' OR COLUMN_NAME LIKE '%%ADDRESS%%' OR COLUMN_NAME LIKE '%%TANTOU%%' OR COLUMN_NAME LIKE '%%CREATE_USER%%' OR COLUMN_NAME LIKE '%%UPDATE_USER%%' OR COLUMN_NAME LIKE '%%FURIGANA%%') AND COLUMN_NAME NOT LIKE '%%CD%%')" -h -1 -W -s "," > "%TEMP_COLUMNS%"
        
        REM Check if columns exist
        set "COLUMN_COUNT=0"
        for /f %%C in (%TEMP_COLUMNS%) do set /a COLUMN_COUNT+=1
        
        if !COLUMN_COUNT! gtr 0 (
            REM Build column list
            set "COLUMNS="
            for /f "tokens=*" %%C in (%TEMP_COLUMNS%) do (
                if "!COLUMNS!"=="" (
                    set "COLUMNS=%%C"
                ) else (
                    set "COLUMNS=!COLUMNS!, %%C"
                )
            )
            
            REM Export data
            sqlcmd -S "%CONN_STR%" -Q "SELECT !COLUMNS! FROM !TABLE_NAME!" -h -1 -W -s "	" > "%TEMP_DATA%"
            
            REM Create UPDATE script
            set "SCRIPT_FILE=%OUTPUT_DIR%\!TABLE_NAME!_update.sql"
            echo -- UPDATE script for table !TABLE_NAME! ^(columns containing NAME, TEL, FAX, POST, ADDRESS, TANTOU, CREATE_USER, UPDATE_USER, FURIGANA^) > "!SCRIPT_FILE!"
            echo. >> "!SCRIPT_FILE!"
            
            set "ROW_NUM=0"
            for /f "tokens=*" %%D in (%TEMP_DATA%) do (
                set /a ROW_NUM+=1
                set "DATA_ROW=%%D"
                
                REM Build SET clause (simplified - would need more complex parsing for real implementation)
                echo ;WITH T AS ^(SELECT *, ROW_NUMBER^(^) OVER ^(ORDER BY ^(SELECT 1^)^) AS rn FROM !TABLE_NAME!^) >> "!SCRIPT_FILE!"
                echo UPDATE T SET [COLUMN_UPDATES_HERE] WHERE rn = !ROW_NUM!; >> "!SCRIPT_FILE!"
                echo. >> "!SCRIPT_FILE!"
            )
            
            echo Created UPDATE script: !SCRIPT_FILE!
        ) else (
            echo Table !TABLE_NAME! has no target columns
        )
    )
)

REM Cleanup
del "%TEMP_TABLES%" 2>nul
del "%TEMP_COLUMNS%" 2>nul
del "%TEMP_DATA%" 2>nul

echo Script execution completed.
pause