# SOUA.ps1
# ---
# This script assists in installing Smart Office.
# It ensures necessary prerequisites are met, processes are managed, and services are configured.
# ---
# Version 1.31
# - Included handling of SMUpdates.exe process at appropriate steps
# - Added success/failure messages at the end of each part
# - Updated user messages for professionalism and consistency

# Part 1 - Check for Admin Rights
# -----
Write-Host "[Part 1/11] Checking for admin rights..."
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Admin)) {
    Write-Host "Script is not running with admin rights. Please run as administrator." -ForegroundColor Red
    pause
    exit
} else {
    Write-Host "Admin rights confirmed." -ForegroundColor Green
}

# Part 2 - Check for Running Smart Office Processes
# -----
Write-Host "[Part 2/11] Checking for running Smart Office processes..."
$processesToCheck = @("Sm32Main", "Sm32")

foreach ($process in $processesToCheck) {
    if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
        Write-Host "$process is running. Please close it before proceeding." -ForegroundColor Red
        pause
        exit
    } else {
        Write-Host "$process is not running." -ForegroundColor Green
    }
}

# Part 3 - Create Directory if it Doesn't Exist
# -----
Write-Host "[Part 3/11] Ensuring working directory exists..."
$workingDir = "C:\winsm"

if (-not (Test-Path $workingDir)) {
    New-Item -Path $workingDir -ItemType Directory | Out-Null
    Write-Host "Directory $workingDir created." -ForegroundColor Green
} else {
    Write-Host "Directory $workingDir already exists." -ForegroundColor Green
}

# Part 4 - Download and Run SO_UC.exe
# -----
Write-Host "[Part 4/11] Downloading and running SO_UC.exe if necessary..."
$SO_UC_Path = "$workingDir\SO_UC.exe"
$SO_UC_URL = "https://github.com/SMControl/SO_UC/blob/main/SO_UC.exe"

if (-not (Test-Path $SO_UC_Path)) {
    Write-Host "SO_UC.exe not found. Downloading..."
    Invoke-WebRequest -Uri $SO_UC_URL -OutFile $SO_UC_Path
    Write-Host "Download completed." -ForegroundColor Green
} else {
    Write-Host "SO_UC.exe already exists. Skipping download." -ForegroundColor Green
}

Start-Process -FilePath $SO_UC_Path -Wait
Write-Host "SO_UC.exe execution completed." -ForegroundColor Green

# Part 5 - Check for Firebird Installation
# -----
Write-Host "[Part 5/11] Checking for Firebird installation..."
$firebirdDir = "C:\Program Files (x86)\Firebird"
$firebirdInstallerURL = "https://raw.githubusercontent.com/SMControl/SM_Firebird_Installer/main/SMFI_Online.ps1"

if (-not (Test-Path $firebirdDir)) {
    Write-Host "Firebird not found. Installing..."
    Invoke-Expression -Command (irm $firebirdInstallerURL | iex)
    Write-Host "Firebird installation completed." -ForegroundColor Green
} else {
    Write-Host "Firebird already installed." -ForegroundColor Green
}

# Part 5.1 - Stop SMUpdates.exe if running
# -----
Write-Host "[Part 5.1/11] Checking and stopping SMUpdates.exe if running..."
Stop-Process -Name "SMUpdates" -ErrorAction SilentlyContinue
Write-Host "SMUpdates.exe process terminated if it was running." -ForegroundColor Green

# Part 6 - Check and Manage Smart Office Live Sales Service
# -----
Write-Host "[Part 6/11] Checking and managing Smart Office Live Sales service..."
$ServiceName = "srvSOLiveSales"
$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue

$initialServiceState = $null
if ($service) {
    $initialServiceState = $service.Status
    if ($service.Status -eq 'Running') {
        Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
        Set-Service -Name $ServiceName -StartupType Disabled
        Write-Host "Stopped and disabled $ServiceName service." -ForegroundColor Green
    } else {
        Write-Host "$ServiceName service is not running. No action needed." -ForegroundColor Green
    }
} else {
    Write-Host "$ServiceName service does not exist. No action needed." -ForegroundColor Green
}

# Part 7 - Check and Manage PDTWiFi Processes
# -----
Write-Host "[Part 7/11] Checking and managing PDTWiFi processes..."
$PDTWiFiProcesses = @("PDTWiFi", "PDTWiFi64")
$initialPDTWiFiStates = @{}

foreach ($process in $PDTWiFiProcesses) {
    $p = Get-Process -Name $process -ErrorAction SilentlyContinue
    if ($p) {
        $initialPDTWiFiStates[$process] = $true
        Stop-Process -Name $process -Force -ErrorAction SilentlyContinue
        Write-Host "Stopped $process process." -ForegroundColor Green
    } else {
        $initialPDTWiFiStates[$process] = $false
        Write-Host "$process process is not running. No action needed." -ForegroundColor Green
    }
}

# Part 8 - Launch Setup Executable
# -----
Write-Host "[Part 8/11] Launching Smart Office setup executable..."
$setupDir = "$workingDir\SmartOffice_Installer"
$setupExe = Get-ChildItem -Path $setupDir -Filter "Setup*.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if ($setupExe) {
    Start-Process -FilePath $setupExe.FullName -Wait
    Write-Host "Smart Office setup completed." -ForegroundColor Green
} else {
    Write-Host "Setup executable not found in $setupDir." -ForegroundColor Red
    pause
    exit
}

# Part 9 - Wait for User Confirmation
# -----
Write-Host "[Part 9/11] Please press Enter when the Smart Office installation is fully finished..."
Read-Host "When the installation of Smart Office is fully finished, please press Enter to finish off installation assistant tasks"

# Part 10 - Set Permissions for StationMaster Folder
# -----
Write-Host "[Part 10/11] Setting permissions for StationMaster folder..."

# Kill SMUpdates.exe if running again
Stop-Process -Name "SMUpdates" -ErrorAction SilentlyContinue
Write-Host "SMUpdates.exe process terminated if it was running."

& icacls "C:\Program Files (x86)\StationMaster" /grant "*S-1-1-0:(OI)(CI)F" /T /C > $null
if ($LASTEXITCODE -eq 0) {
    Write-Host "Permissions for StationMaster folder set successfully." -ForegroundColor Green
} else {
    Write-Host "Failed to set permissions for StationMaster folder." -ForegroundColor Red
}

# Part 11 - Revert Services and Processes to Original State
# -----
Write-Host "[Part 11/11] Reverting services and processes to original state..."

# Revert srvSOLiveSales service
if ($initialServiceState -eq 'Running') {
    Set-Service -Name $ServiceName -StartupType Automatic
    Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
    Write-Host "$ServiceName service re-enabled and started successfully." -ForegroundColor Green
} elseif ($initialServiceState -eq 'Stopped') {
    Set-Service -Name $ServiceName -StartupType Manual
    Write-Host "$ServiceName service set back to manual startup." -ForegroundColor Green
}

# Revert PDTWiFi processes
foreach ($process in $PDTWiFiProcesses) {
    if ($initialPDTWiFiStates[$process]) {
        Start-Process -FilePath "C:\Program Files (x86)\StationMaster\$process.exe"
        Write-Host "Started $process process." -ForegroundColor Green
    }
}

Write-Host "All tasks completed successfully." -ForegroundColor Green
Write-Host "Press any key to exit..."
$null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
