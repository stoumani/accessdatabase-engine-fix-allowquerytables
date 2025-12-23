# this script requires Administrator prvileges
# To fix the issue that is occurring for Excel Pivot Tables since update KB5072033 for Windows, adds a registry value that allows query remote tables.
# Link to the issue and information found is located here https://learn.microsoft.com/en-us/answers/questions/5658888/excel-pivot-tables-linked-to-access-not-refreshing

# check if running as Admin, if not, error!
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Please right-click and select 'Run as Administrator'" -ForegroundColor Yellow
    pause
    exit 1
}

# registry path and value that will be added
$registryPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Access Connectivity Engine\Engines"
$valueName = "AllowQueryRemoteTables"
$valueData = 1
$valueType = "DWord"

try {
    # check if it exists, if not, it will create
    if (-not (Test-Path $registryPath)) {
        Write-Host "Registry path does not exist. Creating path..." -ForegroundColor Yellow
        New-Item -Path $registryPath -Force | Out-Null
        Write-Host "Registry path created successfully." -ForegroundColor Green
    }
    
    # set registry value
    Write-Host "Setting registry value..." -ForegroundColor Cyan
    Set-ItemProperty -Path $registryPath -Name $valueName -Value $valueData -Type $valueType -Force
    
    # verify
    $currentValue = Get-ItemProperty -Path $registryPath -Name $valueName -ErrorAction SilentlyContinue
    
    if ($currentValue.$valueName -eq $valueData) {
        Write-Host "SUCCESS: Registry value set successfully!" -ForegroundColor Green
        Write-Host "Path: $registryPath" -ForegroundColor Cyan
        Write-Host "Value: $valueName = $valueData" -ForegroundColor Cyan
    } else {
        Write-Host "WARNING: Unable to verify the registry value was set correctly." -ForegroundColor Yellow
    }
    
} catch {
    Write-Host "ERROR: Failed to set registry value." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

Write-Host "`nPress any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
