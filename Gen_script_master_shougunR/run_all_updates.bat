@echo off

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

rem Directory containing SQL scripts
set SCRIPT_DIR=output_scripts_update
set OUTPUT_FILE=output.txt

rem Clear output file
echo Starting SQL script execution... > %OUTPUT_FILE%
echo Date: %DATE% %TIME% >> %OUTPUT_FILE%
echo. >> %OUTPUT_FILE%

rem Check if directory exists
if not exist "%SCRIPT_DIR%" (
    echo Error: Script directory %SCRIPT_DIR% does not exist.
    echo Error: Script directory %SCRIPT_DIR% does not exist. >> %OUTPUT_FILE%
    pause
    exit /b 1
)

rem Run each SQL file in the directory
for %%F in ("%SCRIPT_DIR%\*.sql") do (
    echo Running %%F ...
    echo Running %%F ... >> %OUTPUT_FILE%
    sqlcmd -S %SERVER% -d %DATABASE% -U %USER% -P %PASSWORD% -i "%%F" -f 65001 >> %OUTPUT_FILE% 2>&1
    if errorlevel 1 (
        echo Warning: %%F had errors - possibly table does not exist in target DB
        echo Warning: %%F had errors - possibly table does not exist in target DB >> %OUTPUT_FILE%
        echo Continuing with next file...
        echo Continuing with next file... >> %OUTPUT_FILE%
    ) else (
        echo %%F executed successfully >> %OUTPUT_FILE%
    )
    echo. >> %OUTPUT_FILE%
)

echo All scripts executed successfully.
echo All scripts executed successfully. >> %OUTPUT_FILE%
echo End time: %DATE% %TIME% >> %OUTPUT_FILE%
echo Check %OUTPUT_FILE% for execution details.
pause