# Script_Name: soua.ps1
# ---
# This script assists in installing Smart Office.
# ---
# Version 1.31
# - Improved readability and clarity of comments.
# - Standardized user messages for consistency.
# - Separated functionality into clear parts.
# - Managed service and process states cleanly and consistently.

# Start time for the script
$startTime = Get-Date

# Function to check and manage services
function Manage-Service {
    param (
        [string]$ServiceName,
        [string]$Action
    )
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($null -ne $service) {
        switch ($Action) {
            "Stop" {
                if ($service.Status -eq 'Running') {
                    Write-Host "Stopping $ServiceName service..." -ForegroundColor Yellow
                    Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
                    $service.WaitForStatus('Stopped', '00:01:00')
                    Write-Host "$ServiceName service stopped." -ForegroundColor Green
                }
            }
            "Disable" {
                if ($service.StartType -ne 'Disabled') {
                    Write-Host "Disabling $ServiceName service..." -ForegroundColor Yellow
                    Set-Service -Name $ServiceName -StartupType Disabled
                    Write-Host "$ServiceName service disabled." -ForegroundColor Green
                }
            }
            "Enable" {
                if ($service.StartType -eq 'Disabled') {
                    Write-Host "Enabling $ServiceName service..." -ForegroundColor Yellow
                    Set-Service -Name $ServiceName -StartupType Automatic
                    Write-Host "$ServiceName service enabled." -ForegroundColor Green
                }
            }
            "Start" {
                if ($service.Status -ne 'Running') {
                    Write-Host "Starting $ServiceName service..." -ForegroundColor Yellow
                    Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
                    $service.WaitForStatus('Running', '00:01:00')
                    Write-Host "$ServiceName service started successfully." -ForegroundColor Green
                }
            }
        }
    }
}

# Part 1 - Check for Admin Rights
# -----
Write-Host "[Part 1/11] Checking for administrative rights..." -ForegroundColor Cyan
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script requires administrative rights. Please run as administrator." -ForegroundColor Red
    Read-Host "Press any key to exit..."
    exit
}
Write-Host "[Part 1/11] Check for administrative rights completed successfully." -ForegroundColor Green

# Part 2 - Check for Running Smart Office Processes
# -----
Write-Host "[Part 2/11] Checking for running Smart Office processes..." -ForegroundColor Cyan
$smartOfficeProcesses = @("Sm32Main", "Sm32")
foreach ($process in $smartOfficeProcesses) {
    if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
        Write-Host "Process $process is running. Please close it before proceeding." -ForegroundColor Red
        Read-Host "Press any key to exit..."
        exit
    }
}
Write-Host "[Part 2/11] Check for running Smart Office processes completed successfully." -ForegroundColor Green

# Part 3 - Create Directory if it Doesn't Exist
# -----
Write-Host "[Part 3/11] Creating C:\winsm directory if it doesn't exist..." -ForegroundColor Cyan
if (-not (Test-Path -Path "C:\winsm")) {
    New-Item -Path "C:\winsm" -ItemType Directory | Out-Null
    Write-Host "Directory C:\winsm created." -ForegroundColor Green
} else {
    Write-Host "Directory C:\winsm already exists." -ForegroundColor Yellow
}
Write-Host "[Part 3/11] Create directory completed successfully." -ForegroundColor Green

# Part 4 - Download and Run SO_UC.exe
# -----
Write-Host "[Part 4/11] Checking for SO_UC.exe..." -ForegroundColor Cyan
$soUcExePath = "C:\winsm\SO_UC.exe"
$soUcExeUrl = "https://github.com/SMControl/SO_UC/blob/main/SO_UC.exe"
if (-not (Test-Path -Path $soUcExePath)) {
    try {
        Write-Host "Downloading SO_UC.exe..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $soUcExeUrl -OutFile $soUcExePath
        Write-Host "SO_UC.exe downloaded successfully." -ForegroundColor Green
    } catch {
        Write-Host "Failed to download SO_UC.exe. Please check your internet connection and try again." -ForegroundColor Red
        Read-Host "Press any key to exit..."
        exit
    }
} else {
    Write-Host "SO_UC.exe already exists." -ForegroundColor Yellow
}
Write-Host "Running SO_UC.exe..." -ForegroundColor Yellow
Start-Process -FilePath $soUcExePath -Wait
Write-Host "SO_UC.exe finished running." -ForegroundColor Green
Write-Host "[Part 4/11] SO_UC.exe check and run completed successfully." -ForegroundColor Green

# Part 5 - Check for Firebird Installation
# -----
Write-Host "[Part 5/11] Checking for Firebird installation..." -ForegroundColor Cyan
$firebirdDir = "C:\Program Files (x86)\Firebird"
$firebirdInstallerUrl = "https://raw.githubusercontent.com/SMControl/SM_Firebird_Installer/main/SMFI_Online.ps1"
if (-not (Test-Path -Path $firebirdDir)) {
    Write-Host "Firebird is not installed. Running the online installer..." -ForegroundColor Yellow
    Invoke-Expression -Command "irm $firebirdInstallerUrl | iex"
    Write-Host "Firebird installation completed." -ForegroundColor Green
} else {
    Write-Host "Firebird is already installed." -ForegroundColor Yellow
}
Write-Host "[Part 5/11] Check for Firebird installation completed successfully." -ForegroundColor Green

# Part 6 - Check and Manage Smart Office Live Sales Service
# -----
Write-Host "[Part 6/11] Checking and managing Smart Office Live Sales service..." -ForegroundColor Cyan
$ServiceName = "srvSOLiveSales"
$Service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($Service -ne $null -and $Service.Status -eq 'Running') {
    Manage-Service -ServiceName $ServiceName -Action "Stop"
    Manage-Service -ServiceName $ServiceName -Action "Disable"
}
Write-Host "[Part 6/11] Check and manage Smart Office Live Sales service completed successfully." -ForegroundColor Green

# Part 7 - Check and Manage PDTWiFi Processes
# -----
Write-Host "[Part 7/11] Checking and managing PDTWiFi processes..." -ForegroundColor Cyan
$ProcessesToCheck = @("PDTWiFi", "PDTWiFi64")
$ProcessesClosed = @()
foreach ($ProcessName in $ProcessesToCheck) {
    $Process = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if ($Process) {
        Write-Host "Stopping $ProcessName process..." -ForegroundColor Yellow
        Stop-Process -Name $ProcessName -Force -ErrorAction SilentlyContinue
        $ProcessesClosed += $ProcessName
        Write-Host "$ProcessName process stopped." -ForegroundColor Green
    }
}
Write-Host "[Part 7/11] Check and manage PDTWiFi processes completed successfully." -ForegroundColor Green

# Part 8 - Launch Setup Executable
# -----
Write-Host "[Part 8/11] Launching Smart Office setup executable..." -ForegroundColor Cyan
$setupExePath = Get-ChildItem -Path "C:\winsm\SmartOffice_Installer" -Filter "Setup*.exe" | Select-Object -First 1
if ($setupExePath) {
    Start-Process -FilePath $setupExePath.FullName -Wait
    Write-Host "Smart Office setup executable finished." -ForegroundColor Green
} else {
    Write-Host "Setup executable not found. Please ensure it is located in C:\winsm\SmartOffice_Installer." -ForegroundColor Red
    Read-Host "Press any key to exit..."
    exit
}
Write-Host "[Part 8/11] Launch Smart Office setup executable completed successfully." -ForegroundColor Green

# Part 9 - Wait for User Confirmation
# -----
Write-Host "[Part 9/11] Waiting for user confirmation..." -ForegroundColor Cyan
Read-Host "When installation of Smart Office is fully finished, please press Enter to finish off installation assistant tasks."
Write-Host "[Part 9/11] User confirmation completed successfully." -ForegroundColor Green

# Part 10 - Set Permissions for StationMaster Folder
# -----
Write-Host "[Part 10/11] Setting permissions for StationMaster folder..." -ForegroundColor Cyan
icacls "C:\Program Files (x86)\StationMaster" /grant "*S-1-1-0:(OI)(CI)F" /T /C | Out-Null
Write-Host "Permissions for StationMaster

 folder set." -ForegroundColor Green
Write-Host "[Part 10/11] Set permissions for StationMaster folder completed successfully." -ForegroundColor Green

# Part 11 - Revert Services and Processes to Original State
# -----
Write-Host "[Part 11/11] Reverting services and processes to their original state..." -ForegroundColor Cyan
if ($Service -ne $null -and $Service.Status -eq 'Stopped') {
    Manage-Service -ServiceName $ServiceName -Action "Enable"
    Manage-Service -ServiceName $ServiceName -Action "Start"
}

foreach ($ProcessName in $ProcessesClosed) {
    $ProcessPath = "C:\Program Files (x86)\StationMaster\$ProcessName.exe"
    if (Test-Path -Path $ProcessPath) {
        Write-Host "Starting $ProcessName process..." -ForegroundColor Yellow
        Start-Process -FilePath $ProcessPath
        Write-Host "$ProcessName process started." -ForegroundColor Green
    }
}
Write-Host "[Part 11/11] Revert services and processes to their original state completed successfully." -ForegroundColor Green

# Total script run time
$endTime = Get-Date
$totalTime = $endTime - $startTime
Write-Host "Total script run time: $totalTime" -ForegroundColor Cyan

# End of script
Read-Host "Press any key to exit..."
