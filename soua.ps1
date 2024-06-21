# Script_Name: soua.ps1
# ---
# This script assists in installing Smart Office.
# ---
# Version 1.30
# - Added prompt for user before starting Part 9 to ensure installation is complete.
# - Improved readability and clarity of comments.
# - Added detailed status messages for handling PDTWiFi.exe and PDTWiFi64.exe.
# - Included total run time summary at the end of the script.

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
        if ($Action -eq "Stop" -and $service.Status -eq 'Running') {
            Write-Host "Stopping $ServiceName service..." -ForegroundColor Yellow
            Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
            $service.WaitForStatus('Stopped', '00:01:00')
            Write-Host "$ServiceName service stopped." -ForegroundColor Green
        } elseif ($Action -eq "Disable" -and $service.StartType -ne 'Disabled') {
            Write-Host "Disabling $ServiceName service..." -ForegroundColor Yellow
            Set-Service -Name $ServiceName -StartupType Disabled
            Write-Host "$ServiceName service disabled." -ForegroundColor Green
        } elseif ($Action -eq "Enable" -and $service.StartType -eq 'Disabled') {
            Write-Host "Enabling $ServiceName service..." -ForegroundColor Yellow
            Set-Service -Name $ServiceName -StartupType Automatic
            $service.WaitForStatus('Stopped', '00:01:00') # Ensure service is in stopped state before starting
            Write-Host "$ServiceName service enabled." -ForegroundColor Green
        } elseif ($Action -eq "Start" -and $service.Status -ne 'Running') {
            Write-Host "Starting $ServiceName service..." -ForegroundColor Yellow
            Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
            $service.WaitForStatus('Running', '00:01:00')
            Write-Host "$ServiceName service started successfully." -ForegroundColor Green
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

# Part 3 - Create Directory if it Doesn't Exist
# -----
Write-Host "[Part 3/11] Creating C:\winsm directory if it doesn't exist..." -ForegroundColor Cyan
if (-not (Test-Path -Path "C:\winsm")) {
    New-Item -Path "C:\winsm" -ItemType Directory | Out-Null
}

# Part 4 - Download and Run SO_UC.exe
# -----
Write-Host "[Part 4/11] Downloading and running SO_UC.exe..." -ForegroundColor Cyan
$soUcExePath = "C:\winsm\SO_UC.exe"
$soUcExeUrl = "https://github.com/SMControl/SO_UC/blob/main/SO_UC.exe"
if (-not (Test-Path -Path $soUcExePath)) {
    try {
        Invoke-WebRequest -Uri $soUcExeUrl -OutFile $soUcExePath
        Write-Host "SO_UC.exe downloaded successfully." -ForegroundColor Green
    } catch {
        Write-Host "Failed to download SO_UC.exe. Please check your internet connection and try again." -ForegroundColor Red
        Read-Host "Press any key to exit..."
        exit
    }
}
Start-Process -FilePath $soUcExePath -Wait

# Part 5 - Check for Firebird Installation
# -----
Write-Host "[Part 5/11] Checking for Firebird installation..." -ForegroundColor Cyan
$firebirdDir = "C:\Program Files (x86)\Firebird"
$firebirdInstallerUrl = "https://raw.githubusercontent.com/SMControl/SM_Firebird_Installer/main/SMFI_Online.ps1"
if (-not (Test-Path -Path $firebirdDir)) {
    Write-Host "Firebird is not installed. Running the online installer..." -ForegroundColor Yellow
    Invoke-Expression -Command "irm $firebirdInstallerUrl | iex"
}

# Part 6 - Check and Manage Smart Office Live Sales Service
# -----
Write-Host "[Part 6/11] Checking and managing Smart Office Live Sales service..." -ForegroundColor Cyan
$ServiceName = "srvSOLiveSales"
$Service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($Service -ne $null -and $Service.Status -eq 'Running') {
    Manage-Service -ServiceName $ServiceName -Action "Stop"
    Manage-Service -ServiceName $ServiceName -Action "Disable"
}

# Part 7 - Check and Manage PDTWiFi Processes
# -----
Write-Host "[Part 7/11] Checking and managing PDTWiFi processes..." -ForegroundColor Cyan
$ProcessesToCheck = @("PDTWiFi", "PDTWiFi64")
$ProcessesClosed = @()
foreach ($ProcessName in $ProcessesToCheck) {
    $Process = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if ($Process) {
        Write-Host "Stopping $ProcessName..." -ForegroundColor Yellow
        Stop-Process -Name $ProcessName -Force -ErrorAction SilentlyContinue
        $ProcessesClosed += $ProcessName
        Write-Host "$ProcessName stopped." -ForegroundColor Green
    }
}

# Part 8 - Launch Setup Executable
# -----
Write-Host "[Part 8/11] Launching Smart Office setup executable..." -ForegroundColor Cyan
$setupExePath = Get-ChildItem -Path "C:\winsm\SmartOffice_Installer" -Filter "Setup*.exe" | Select-Object -First 1
if ($setupExePath) {
    Start-Process -FilePath $setupExePath.FullName -Wait
} else {
    Write-Host "Setup executable not found. Please ensure it is located in C:\winsm\SmartOffice_Installer." -ForegroundColor Red
    Read-Host "Press any key to exit..."
    exit
}

# Part 9 - Wait for User Confirmation
# -----
Write-Host "[Part 9/11] Ensuring Smart Office processes are closed..." -ForegroundColor Cyan
Read-Host "Press any key to continue after installation is fully finished..."

# Part 10 - Set Permissions for StationMaster Folder
# -----
Write-Host "[Part 10/11] Setting permissions for StationMaster folder..." -ForegroundColor Cyan
# Set permissions for StationMaster folder
icacls "C:\Program Files (x86)\StationMaster" /grant "*S-1-1-0:(OI)(CI)F" /T /C | Out-Null

# Part 11 - Revert Changes
# -----
Write-Host "[Part 11/11] Reverting changes..." -ForegroundColor Cyan

# Restart srvSOLiveSales service if it was running
If ($Service -ne $null -and $Service.Status -eq 'Stopped') {
    Manage-Service -ServiceName $ServiceName -Action "Enable"
    Manage-Service -ServiceName $ServiceName -Action "Start"
}

# Restart PDTWiFi.exe and PDTWiFi64.exe if they were previously running
$ProcessesToRestart = @("PDTWiFi", "PDTWiFi64")
ForEach ($ProcessName in $ProcessesToRestart) {
    If ($ProcessesClosed -contains $ProcessName) {
        Write-Host "Starting $ProcessName..." -ForegroundColor Yellow
        Start-Process -FilePath "C:\Program Files (x86)\StationMaster\$ProcessName.exe" -ErrorAction SilentlyContinue
        Write-Host "$ProcessName started." -ForegroundColor Green
    }
}

# Initialize end time
$endTime = Get-Date

# Calculate total script run time
$totalTime = New-TimeSpan -Start $startTime -End $endTime

# Display total script run time
Write-Host "Script completed in $($totalTime.ToString('hh\:mm\:ss'))" -ForegroundColor Green

Read-Host "Press any key to exit..."
