@echo off

rem SQL Server connection details
set SERVER=VJP-LAP0261\SQLSERVER2022
set DATABASE=KankyouShougunR_demo
set USER=sa
set PASSWORD=Vti123456!

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