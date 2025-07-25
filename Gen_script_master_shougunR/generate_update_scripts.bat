@echo off
setlocal EnableDelayedExpansion

rem =====================================================================
rem UPDATE Script Generator (Batch Version)
rem =====================================================================

rem === DATABASE CONNECTION SETTINGS ===
rem User can modify these settings directly
set SERVER=VJP-LAP0261\SQLSERVER2022
set DATABASE=KankyouShougunR_demo
set USER=sa
set PASSWORD=Vti123456!

rem Output directory for update scripts
set OUTPUT_DIR=output_scripts_update
set LOG_FILE=generate_scripts.log

rem Create output directory if not exists
if not exist "%OUTPUT_DIR%" (
    mkdir "%OUTPUT_DIR%"
)

rem Clear log file
echo Starting update script generation... > %LOG_FILE%
echo Date: %DATE% %TIME% >> %LOG_FILE%
echo. >> %LOG_FILE%

rem Debug: echo connection info (without password)
echo SERVER=%SERVER%
echo DATABASE=%DATABASE%
echo USER=%USER%
echo Password set: [Protected]
echo.

echo SERVER=%SERVER% >> %LOG_FILE%
echo DATABASE=%DATABASE% >> %LOG_FILE%
echo USER=%USER% >> %LOG_FILE%
echo. >> %LOG_FILE%

echo Generating UPDATE scripts from database %DATABASE%...
echo Connecting to %SERVER%...

rem Test connection first
echo Testing database connection...
setlocal DisableDelayedExpansion
sqlcmd -S "%SERVER%" -d "%DATABASE%" -U "%USER%" -P "%PASSWORD%" -Q "SELECT 1 as test" -h -1 2>connection_test.log
setlocal EnableDelayedExpansion

if errorlevel 1 (
    echo Error: Could not connect to database
    echo Connection error details:
    type connection_test.log
    echo.
    echo Error: Could not connect to database >> %LOG_FILE%
    echo Check connection parameters in connect_string.txt >> %LOG_FILE%
    pause
    exit /b 1
) else (
    echo Database connection successful!
    echo Database connection successful! >> %LOG_FILE%
)

rem Get list of tables starting with M_
echo Getting list of tables...
echo Getting list of tables... >> %LOG_FILE%

setlocal DisableDelayedExpansion
sqlcmd -S "%SERVER%" -d "%DATABASE%" -U "%USER%" -P "%PASSWORD%" -Q "SET NOCOUNT ON; SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME LIKE 'M_%%'" -h -1 -W > temp_tables.txt 2>> %LOG_FILE%
setlocal EnableDelayedExpansion

if errorlevel 1 (
    echo Error: Could not get table list
    echo Error: Could not get table list >> %LOG_FILE%
    pause
    exit /b 1
)

rem Check if any tables found
set TABLE_COUNT=0
echo Debug: Checking temp_tables.txt content:
type temp_tables.txt
echo.
echo Debug: End of temp_tables.txt
echo.

for /f "skip=1 tokens=*" %%T in (temp_tables.txt) do (
    if not "%%T"=="" (
        set /a TABLE_COUNT+=1
        echo Debug: Found table: %%T
    )
)

if %TABLE_COUNT%==0 (
    echo No tables starting with M_ found in database
    echo No tables starting with M_ found in database >> %LOG_FILE%
    pause
    exit /b 1
)

echo Found %TABLE_COUNT% tables to process
echo Found %TABLE_COUNT% tables to process >> %LOG_FILE%

rem Process each table
for /f "skip=1 tokens=*" %%T in (temp_tables.txt) do (
    set TABLE_NAME=%%T
    if not "%%T"=="" (
        call :ProcessTable "%%T"
    )
)

rem Clean up
del temp_tables.txt 2>nul
del connection_test.log 2>nul

echo.
echo All UPDATE scripts generated successfully.
echo All UPDATE scripts generated successfully. >> %LOG_FILE%
echo End time: %DATE% %TIME% >> %LOG_FILE%
echo Check %OUTPUT_DIR% folder for generated scripts.
pause
goto :eof

:ProcessTable
set TABLE_NAME=%~1
if "%TABLE_NAME%"=="" goto :eof

echo Processing table %TABLE_NAME%...
echo Processing table %TABLE_NAME%... >> %LOG_FILE%

rem Get columns containing target keywords
setlocal DisableDelayedExpansion
sqlcmd -S "%SERVER%" -d "%DATABASE%" -U "%USER%" -P "%PASSWORD%" -Q "SET NOCOUNT ON; SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '%TABLE_NAME%' AND ((COLUMN_NAME LIKE '%%NAME%%' OR COLUMN_NAME LIKE '%%TEL%%' OR COLUMN_NAME LIKE '%%FAX%%' OR COLUMN_NAME LIKE '%%POST%%' OR COLUMN_NAME LIKE '%%ADDRESS%%' OR COLUMN_NAME LIKE '%%TANTOU%%' OR COLUMN_NAME LIKE '%%CREATE_USER%%' OR COLUMN_NAME LIKE '%%UPDATE_USER%%' OR COLUMN_NAME LIKE '%%FURIGANA%%') AND COLUMN_NAME NOT LIKE '%%CD%%')" -h -1 -W > temp_columns_%TABLE_NAME%.txt 2>> %LOG_FILE%
setlocal EnableDelayedExpansion

if errorlevel 1 (
    echo Warning: Could not get columns for table %TABLE_NAME%
    echo Warning: Could not get columns for table %TABLE_NAME% >> %LOG_FILE%
    goto :ProcessTableEnd
)

rem Check if any target columns exist
set HAS_TARGET_COLUMNS=0
for /f "tokens=*" %%C in (temp_columns_%TABLE_NAME%.txt) do (
    if not "%%C"=="" (
        set HAS_TARGET_COLUMNS=1
        goto :CheckColumnsEnd
    )
)
:CheckColumnsEnd

if %HAS_TARGET_COLUMNS%==0 (
    echo Table %TABLE_NAME% has no target columns
    echo Table %TABLE_NAME% has no target columns >> %LOG_FILE%
    goto :ProcessTableEnd
)

rem Build column list
set COLUMN_LIST=
for /f "tokens=*" %%C in (temp_columns_%TABLE_NAME%.txt) do (
    if not "%%C"=="" (
        if "!COLUMN_LIST!"=="" (
            set COLUMN_LIST=%%C
        ) else (
            set COLUMN_LIST=!COLUMN_LIST!, %%C
        )
    )
)

if "!COLUMN_LIST!"=="" (
    echo No valid columns found for table %TABLE_NAME%
    echo No valid columns found for table %TABLE_NAME% >> %LOG_FILE%
    goto :ProcessTableEnd
)

echo Found columns: !COLUMN_LIST!
echo Found columns: !COLUMN_LIST! >> %LOG_FILE%

rem Generate UPDATE script header
set SCRIPT_FILE=%OUTPUT_DIR%\%TABLE_NAME%_update.sql
echo -- UPDATE script for table %TABLE_NAME% > "%SCRIPT_FILE%"
echo -- Columns: !COLUMN_LIST! >> "%SCRIPT_FILE%"
echo -- Generated on: %DATE% %TIME% >> "%SCRIPT_FILE%"
echo. >> "%SCRIPT_FILE%"

rem Get row count
setlocal DisableDelayedExpansion
sqlcmd -S "%SERVER%" -d "%DATABASE%" -U "%USER%" -P "%PASSWORD%" -Q "SET NOCOUNT ON; SELECT COUNT(*) FROM %TABLE_NAME%" -h -1 > temp_count.txt 2>> %LOG_FILE%
setlocal EnableDelayedExpansion
set /p ROW_COUNT=<temp_count.txt
echo -- Total rows: %ROW_COUNT% >> "%SCRIPT_FILE%"
echo. >> "%SCRIPT_FILE%"

rem Note about batch limitations
echo -- NOTE: This is a simplified batch-generated script. >> "%SCRIPT_FILE%"
echo -- For production use, please use the Python version for better data handling. >> "%SCRIPT_FILE%"
echo -- Manual review and testing is recommended before execution. >> "%SCRIPT_FILE%"
echo. >> "%SCRIPT_FILE%"

echo Generated %SCRIPT_FILE% with column definitions
echo Generated %SCRIPT_FILE% >> %LOG_FILE%

:ProcessTableEnd
del temp_columns_%TABLE_NAME%.txt 2>nul
del temp_count.txt 2>nul
goto :eof