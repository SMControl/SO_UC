Write-Host "SOUA.ps1" -ForegroundColor Green
# ---
# This script assists in installing Smart Office.
# It ensures necessary prerequisites are met, processes are managed, and services are configured.
# ---
Write-Host "Version 1.102" -ForegroundColor Green
# - removed flag file handling and resume capability



Write-Host "---" -ForegroundColor Green

# Part 1 - Check for Admin Rights
# -----
Write-Host "[Part 1/13] Checking for admin rights..." -ForegroundColor Green
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Admin)) {
    Write-Host "Error: Administrator rights required to run this script. Exiting." -ForegroundColor Red
    pause
    exit
}

# Part 2 - Check for Running Smart Office Processes
# -----
Write-Host "[Part 2/13] Checking for running Smart Office processes..." -ForegroundColor Green
$processesToCheck = @("Sm32Main", "Sm32")
foreach ($process in $processesToCheck) {
    if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
        Write-Host "Error: Smart Office process '$process' is running. Close it and retry." -ForegroundColor Red
        pause
        exit
    }
}

# Part 4 - Download and Run SO_UC.exe Hidden if Necessary
# -----
Write-Host "[Part 4/13] Downloading latest Smart Office Setup if necessary..." -ForegroundColor Green

# Display message about firewall
Write-Host "[WARNING] Please ensure SO_UC.exe is allowed through the firewall." -ForegroundColor Cyan
Write-Host "[WARNING] It's responsible for retrieving the latest Smart Office Setup." -ForegroundColor Cyan

$SO_UC_Path = "$workingDir\SO_UC.exe"
$SO_UC_URL = "https://github.com/SMControl/SO_UC/raw/main/SO_UC.exe"
if (-not (Test-Path $SO_UC_Path)) {
    try {
        Invoke-WebRequest -Uri $SO_UC_URL -OutFile $SO_UC_Path
    } catch {
        Write-Host "Error downloading SO_UC.exe: $_" -ForegroundColor Red
        exit
    }
}

# Start SO_UC.exe hidden
$processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
$processStartInfo.FileName = $SO_UC_Path
$processStartInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

try {
    Start-Process -FilePath $processStartInfo.FileName -Wait -WindowStyle $processStartInfo.WindowStyle
} catch {
    Write-Host "Error starting SO_UC.exe: $_" -ForegroundColor Red
    exit
}

# Part 5 - Check for Firebird Installation
# -----
Write-Host "[Part 5/13] Checking for Firebird installation..." -ForegroundColor Green
$firebirdDir = "C:\Program Files (x86)\Firebird"
$firebirdInstallerURL = "https://raw.githubusercontent.com/SMControl/SM_Firebird_Installer/main/SMFI_Online.ps1"
if (-not (Test-Path $firebirdDir)) {
    try {
        Invoke-Expression -Command (irm $firebirdInstallerURL | iex)
    } catch {
        Write-Host "Error installing Firebird: $_" -ForegroundColor Red
        exit
    }
}

# Part 6 - Stop SMUpdates.exe if Running
# -----
Write-Host "[Part 6/13] Checking and stopping SMUpdates.exe if running..." -ForegroundColor Green
try {
    Stop-Process -Name "SMUpdates" -ErrorAction SilentlyContinue
} catch {
    Write-Host "Error stopping SMUpdates.exe: $_" -ForegroundColor Red
    exit
}

# Part 7 - Check and Manage Smart Office Live Sales Service
# -----
Write-Host "[Part 7/13] Checking and managing Smart Office Live Sales service..." -ForegroundColor Green
$ServiceName = "srvSOLiveSales"
$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue

if ($service -and $service.Status -eq 'Running') {
    try {
        Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
        Set-Service -Name $ServiceName -StartupType Disabled
    } catch {
        Write-Host "Error stopping service '$ServiceName': $_" -ForegroundColor Red
        exit
    }
}

# Part 8 - Check and Manage PDTWiFi Processes
# -----
Write-Host "[Part 8/13] Checking and managing PDTWiFi processes..." -ForegroundColor Green
$PDTWiFiProcesses = @("PDTWiFi", "PDTWiFi64")
foreach ($process in $PDTWiFiProcesses) {
    try {
        $p = Get-Process -Name $process -ErrorAction SilentlyContinue
        if ($p) {
            Stop-Process -Name $process -Force -ErrorAction SilentlyContinue
        }
    } catch {
        Write-Host "Error managing process '$process': $_" -ForegroundColor Red
        exit
    }
}

# Part 9 - Launch Setup Executable
# -----
Write-Host "[Part 9/13] Launching Smart Office setup executable..." -ForegroundColor Green
$setupDir = "C:\Path\To\Your\SmartOffice_Installer"  # Adjust this path to match your setup directory
$setupExe = Get-ChildItem -Path $setupDir -Filter "Setup*.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if ($setupExe) {
    Start-Process -FilePath $setupExe.FullName -Wait
} else {
    Write-Host "Error: Smart Office setup executable not found in '$setupDir'." -ForegroundColor Red
    exit
}

# Part 10 - Wait for User Confirmation
# -----
Write-Host "[Part 10/13] Post Installation" -ForegroundColor Green
Write-Host "[Part 10/13] Please press Enter when the Smart Office installation is FULLY finished..." -ForegroundColor White
Read-Host

# Check for Running Smart Office Processes Again
$processesToCheck = @("Sm32Main", "Sm32")
foreach ($process in $processesToCheck) {
    if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
        Write-Host "Smart Office is still running. Please close it and press Enter to continue..." -ForegroundColor Red
        Read-Host
    }
}

# Part 11 - Set Permissions for StationMaster Folder
# -----

# Silently kill processes PDTWiFi and SMUpdates if running
$processesToKill = @("PDTWiFi", "PDTWiFi64", "SMUpdates")
foreach ($process in $processesToKill) {
    Stop-Process -Name $process -Force -ErrorAction SilentlyContinue
}

Write-Host "[Part 11/13] Setting permissions for StationMaster folder..." -ForegroundColor Green
try {
    & icacls "C:\Program Files (x86)\StationMaster" /grant "*S-1-1-0:(OI)(CI)F" /T /C > $null
} catch {
    Write-Host "Error setting permissions for StationMaster folder: $_" -ForegroundColor Red
}

# Part 12 - Set Permissions for Firebird Folder
# -----
Write-Host "[Part 12/13] Setting permissions for Firebird folder..." -ForegroundColor Green
try {
    & icacls "C:\Program Files (x86)\Firebird" /grant "*S-1-1-0:(OI)(CI)F" /T /C > $null
} catch {
    Write-Host "Error setting permissions for Firebird folder: $_" -ForegroundColor Red
}

# Part 13 - Clean Up and Finish
# -----
Write-Host "[Part 13/13] Cleaning up and finishing script..." -ForegroundColor Green

Write-Host " "

# Calculate and display script execution time
$endTime = Get-Date
$executionTime = $endTime - $startTime
$totalMinutes = [math]::Floor($executionTime.TotalMinutes)
$totalSeconds = $executionTime.Seconds
Write-Host "Script completed successfully in $($totalMinutes)m $($totalSeconds)s." -ForegroundColor Green
