@echo off
REM Batch Script to set AllowQueryRemoteTables registry value
REM This script must be run with Administrator privileges

echo ========================================
echo Registry Value Configuration Script
echo ========================================
echo.

REM Check for Administrator privileges
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

REM Define the registry path and value
set "RegPath=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Access Connectivity Engine\Engines"
set "ValueName=AllowQueryRemoteTables"
set "ValueData=1"
set "ValueType=REG_DWORD"

echo Setting registry value...
echo Path: %RegPath%
echo Value: %ValueName% = %ValueData%
echo.

REM Add/Update the registry value
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