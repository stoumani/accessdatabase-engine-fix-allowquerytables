@echo off
REM The script NEEDS to be run as Admin!
REM To fix the issue that is occurring for Excel Pivot Tables since update KB5072033 for Windows, adds a registry value that allows query remote tables.
REM Link to the issue and information found is located here https://learn.microsoft.com/en-us/answers/questions/5658888/excel-pivot-tables-linked-to-access-not-refreshing


REM check for admin privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ERROR: This script must be run as Administrator!
    echo Please right-click the file and select "Run as administrator" or Run Command Prompt as Administrator, cd into the directory, and execute the file through the terminal!
    echo.
    pause
    exit /b 1
)

echo Running with Administrator privileges...
echo.

REM regsitry path and value found in Microsoft Article
set "RegPath=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Access Connectivity Engine\Engines"
set "ValueName=AllowQueryRemoteTables"
set "ValueData=1"
set "ValueType=REG_DWORD"

echo Setting registry value...
echo Path: %RegPath%
echo Value: %ValueName% = %ValueData%
echo.

REM update registry, if not there, added
reg add "%RegPath%" /v "%ValueName%" /t %ValueType% /d %ValueData% /f

if %errorLevel% equ 0 (
    echo.
    echo SUCCESS: Registry value set successfully, if it doesn't work, restart!
    echo.
) else (
    echo.
    echo ERROR: Failed to set registry value. Make sure the folder directory exists by installing 32-bit Access Database Engine!
    echo Error code: %errorLevel%
    echo.
    pause
    exit /b 1
)

echo Press any key to exit...
pause >nul