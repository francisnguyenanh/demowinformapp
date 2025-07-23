@echo off
setlocal EnableDelayedExpansion

rem =====================================================================
rem UPDATE Script Generator (Batch Version)
rem =====================================================================
rem This script generates UPDATE SQL scripts for tables starting with M_
rem and containing columns with keywords: NAME, TEL, FAX, POST, ADDRESS, 
rem TANTOU, CREATE_USER, UPDATE_USER, FURIGANA
rem 
rem NOTE: This batch version has limitations in data parsing compared to 
rem the Python version (export_update_script.py). For complex data types
rem and better error handling, use the Python script instead.
rem =====================================================================



rem === Read connection info from connect_string.txt using PowerShell ===
powershell -Command "& { $conn = Get-Content 'connect_string.txt'; $conn -replace '.*SERVER=([^;]+).*','SERVER=$1' | Out-File temp_server.txt -Encoding ASCII; $conn -replace '.*DATABASE=([^;]+).*','DATABASE=$1' | Out-File temp_database.txt -Encoding ASCII; $conn -replace '.*UID=([^;]+).*','USER=$1' | Out-File temp_user.txt -Encoding ASCII; $conn -replace '.*PWD=([^;]+).*','PASSWORD=$1' | Out-File temp_password.txt -Encoding ASCII }"

for /f "usebackq tokens=1,2 delims==" %%A in ("temp_server.txt") do set SERVER=%%B
for /f "usebackq tokens=1,2 delims==" %%A in ("temp_database.txt") do set DATABASE=%%B  
for /f "usebackq tokens=1,2 delims==" %%A in ("temp_user.txt") do set USER=%%B
for /f "usebackq tokens=1,2 delims==" %%A in ("temp_password.txt") do set "PASSWORD=%%B"

rem Clean up temp files
del temp_server.txt temp_database.txt temp_user.txt temp_password.txt 2>nul

rem Debug: echo loaded connection info
echo SERVER=%SERVER% DATABASE=%DATABASE% USER=%USER% PASSWORD=%PASSWORD%
echo Password length check: [%PASSWORD%]

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

echo Generating UPDATE scripts from database %DATABASE%...
echo Connecting to %SERVER%...

rem Get list of tables starting with M_
echo Getting list of tables... >> %LOG_FILE%
sqlcmd -S "%SERVER%" -d "%DATABASE%" -U "%USER%" -P "%PASSWORD%" -Q "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME LIKE 'M_%%'" -h -1 -s "," -W -f 65001 > temp_tables.txt 2>> %LOG_FILE%

if errorlevel 1 (
    echo Error: Could not connect to database or get table list
    echo Error: Could not connect to database or get table list >> %LOG_FILE%
    pause
    exit /b 1
)

rem Process each table
for /f "tokens=*" %%T in (temp_tables.txt) do (
    set TABLE_NAME=%%T
    call :ProcessTable "%%T"
)

rem Clean up
del temp_tables.txt 2>nul

echo All UPDATE scripts generated successfully.
echo All UPDATE scripts generated successfully. >> %LOG_FILE%
echo End time: %DATE% %TIME% >> %LOG_FILE%
echo Check %OUTPUT_DIR% folder for generated scripts.
pause
goto :eof

:ProcessTable
set TABLE_NAME=%~1
if "%TABLE_NAME%"=="" goto :eof
if "%TABLE_NAME%"=="TABLE_NAME" goto :eof

echo Processing table %TABLE_NAME%...
echo Processing table %TABLE_NAME%... >> %LOG_FILE%


rem Get columns containing target keywords (NAME, TEL, FAX, POST, ADDRESS, TANTOU, CREATE_USER, UPDATE_USER, FURIGANA) and NOT containing 'CD'
sqlcmd -S "%SERVER%" -d "%DATABASE%" -U "%USER%" -P "%PASSWORD%" -Q "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '%TABLE_NAME%' AND ((COLUMN_NAME LIKE '%%NAME%%' OR COLUMN_NAME LIKE '%%TEL%%' OR COLUMN_NAME LIKE '%%FAX%%' OR COLUMN_NAME LIKE '%%POST%%' OR COLUMN_NAME LIKE '%%ADDRESS%%' OR COLUMN_NAME LIKE '%%TANTOU%%' OR COLUMN_NAME LIKE '%%CREATE_USER%%' OR COLUMN_NAME LIKE '%%UPDATE_USER%%' OR COLUMN_NAME LIKE '%%FURIGANA%%') AND COLUMN_NAME NOT LIKE '%%CD%%')" -h -1 -s "," -W -f 65001 > temp_columns.txt 2>> %LOG_FILE%

if errorlevel 1 (
    echo Warning: Could not get columns for table %TABLE_NAME%
    echo Warning: Could not get columns for table %TABLE_NAME% >> %LOG_FILE%
    goto :ProcessTableEnd
)

rem Check if any target columns exist
set HAS_TARGET_COLUMNS=0
for /f %%C in (temp_columns.txt) do (
    set HAS_TARGET_COLUMNS=1
    goto :CheckColumnsEnd
)
:CheckColumnsEnd

if %HAS_TARGET_COLUMNS%==0 (
    echo Table %TABLE_NAME% has no target columns
    echo Table %TABLE_NAME% has no target columns >> %LOG_FILE%
    goto :ProcessTableEnd
)

rem Build column list
set COLUMN_LIST=
for /f "tokens=*" %%C in (temp_columns.txt) do (
    if "!COLUMN_LIST!"=="" (
        set COLUMN_LIST=%%C
    ) else (
        set COLUMN_LIST=!COLUMN_LIST!, %%C
    )
)

rem Get data from source table
echo Getting data from %TABLE_NAME%... >> %LOG_FILE%
sqlcmd -S "%SERVER%" -d "%DATABASE%" -U "%USER%" -P "%PASSWORD%" -Q "SELECT %COLUMN_LIST% FROM %TABLE_NAME%" -s "	" -W -f 65001 > temp_data.txt 2>> %LOG_FILE%

if errorlevel 1 (
    echo Warning: Could not get data from table %TABLE_NAME%
    echo Warning: Could not get data from table %TABLE_NAME% >> %LOG_FILE%
    goto :ProcessTableEnd
)

rem Generate UPDATE script
set SCRIPT_FILE=%OUTPUT_DIR%\%TABLE_NAME%_update.sql
echo -- UPDATE script for table %TABLE_NAME% ^(columns containing NAME, TEL, FAX, POST, ADDRESS, TANTOU, CREATE_USER, UPDATE_USER, FURIGANA^) > "%SCRIPT_FILE%"
echo. >> "%SCRIPT_FILE%"

set ROW_NUM=1
for /f "skip=2 tokens=*" %%D in (temp_data.txt) do (
    call :GenerateUpdateRow "%%D" !ROW_NUM!
    set /a ROW_NUM+=1
)

echo Generated %SCRIPT_FILE%
echo Generated %SCRIPT_FILE% >> %LOG_FILE%

:ProcessTableEnd
del temp_columns.txt 2>nul
del temp_data.txt 2>nul
goto :eof

:GenerateUpdateRow
set DATA_ROW=%~1
set ROW_NUMBER=%2

if "%DATA_ROW%"=="" goto :eof

rem Parse data row and generate proper SET clauses
rem Note: This is a simplified version. For complex data parsing,
rem the Python script is recommended for better accuracy.
echo ;WITH T AS ^(SELECT *, ROW_NUMBER^(^) OVER ^(ORDER BY ^(SELECT 1^)^) AS rn FROM %TABLE_NAME%^) >> "%SCRIPT_FILE%"

rem For simplicity, we'll generate a basic UPDATE. 
rem In a production environment, proper column-by-column parsing would be needed.
rem The Python version handles this more accurately.
echo UPDATE T SET %COLUMN_LIST% = N'%DATA_ROW%' WHERE rn = %ROW_NUMBER%; >> "%SCRIPT_FILE%"
echo. >> "%SCRIPT_FILE%"

goto :eof
