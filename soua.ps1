# SOUA.ps1
# ---
# This script assists in installing Smart Office.
# It ensures necessary prerequisites are met, processes are managed, and services are configured.
# ---
# Version 1.45
# - Changed message color in Part 10 to white.

# Initialize script start time
$startTime = Get-Date

# Part 1 - Check for Admin Rights
# -----
Write-Host "[Part 1/11] Checking for admin rights..." -ForegroundColor Green
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Admin)) {
    pause
    exit
}

# Part 2 - Check for Running Smart Office Processes
# -----
Write-Host "[Part 2/11] Checking for running Smart Office processes..." -ForegroundColor Green
$processesToCheck = @("Sm32Main", "Sm32")
foreach ($process in $processesToCheck) {
    if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
        pause
        exit
    }
}

# Part 3 - Create Directory if it Doesn't Exist
# -----
Write-Host "[Part 3/11] Ensuring working directory exists..." -ForegroundColor Green
$workingDir = "C:\winsm"
if (-not (Test-Path $workingDir)) {
    New-Item -Path $workingDir -ItemType Directory | Out-Null
}

# Part 4 - Download and Run SO_UC.exe
# -----
Write-Host "[Part 4/11] Downloading latest Smart Office Setup if necessary..." -ForegroundColor Green
$SO_UC_Path = "$workingDir\SO_UC.exe"
$SO_UC_URL = "https://github.com/SMControl/SO_UC/blob/main/SO_UC.exe"
if (-not (Test-Path $SO_UC_Path)) {
    Invoke-WebRequest -Uri $SO_UC_URL -OutFile $SO_UC_Path
}
Start-Process -FilePath $SO_UC_Path -Wait

# Part 5 - Check for Firebird Installation
# -----
Write-Host "[Part 5/11] Checking for Firebird installation..." -ForegroundColor Green
$firebirdDir = "C:\Program Files (x86)\Firebird"
$firebirdInstallerURL = "https://raw.githubusercontent.com/SMControl/SM_Firebird_Installer/main/SMFI_Online.ps1"
if (-not (Test-Path $firebirdDir)) {
    Invoke-Expression -Command (irm $firebirdInstallerURL | iex)
}

# Part 6 - Stop SMUpdates.exe if Running
# -----
Write-Host "[Part 6/11] Checking and stopping SMUpdates.exe if running..." -ForegroundColor Green
Stop-Process -Name "SMUpdates" -ErrorAction SilentlyContinue

# Part 7 - Check and Manage Smart Office Live Sales Service
# -----
Write-Host "[Part 7/11] Checking and managing Smart Office Live Sales service..." -ForegroundColor Green
$ServiceName = "srvSOLiveSales"
$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue

$initialServiceState = $null
if ($service) {
    $initialServiceState = $service.Status
    if ($service.Status -eq 'Running') {
        Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
        Set-Service -Name $ServiceName -StartupType Disabled
    }
}

# Part 8 - Check and Manage PDTWiFi Processes
# -----
Write-Host "[Part 8/11] Checking and managing PDTWiFi processes..." -ForegroundColor Green
$PDTWiFiProcesses = @("PDTWiFi", "PDTWiFi64")
foreach ($process in $PDTWiFiProcesses) {
    $p = Get-Process -Name $process -ErrorAction SilentlyContinue
    if ($p) {
        Stop-Process -Name $process -Force -ErrorAction SilentlyContinue
    }
}

# Part 9 - Launch Setup Executable
# -----
Write-Host "[Part 9/11] Launching Smart Office setup executable..." -ForegroundColor Green
$setupDir = "$workingDir\SmartOffice_Installer"
$setupExe = Get-ChildItem -Path $setupDir -Filter "Setup*.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if ($setupExe) {
    Start-Process -FilePath $setupExe.FullName -Wait
}

# Part 10 - Wait for User Confirmation
# -----
Write-Host "[Part 10/11] Please press Enter when the Smart Office installation is FULLY finished..." -ForegroundColor White
Read-Host

# Part 11 - Set Permissions for StationMaster Folder
# -----
Write-Host "[Part 11/11] Setting permissions for StationMaster folder..." -ForegroundColor Green
Stop-Process -Name "SMUpdates" -ErrorAction SilentlyContinue
& icacls "C:\Program Files (x86)\StationMaster" /grant "*S-1-1-0:(OI)(CI)F" /T /C > $null

# Revert Services and Processes to Original State
# -----
Write-Host "[Revert] Reverting services and processes to original state..." -ForegroundColor Yellow
# Revert srvSOLiveSales service
if ($initialServiceState -eq 'Running') {
    Set-Service -Name $ServiceName -StartupType Automatic
    Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
} elseif ($initialServiceState -eq 'Stopped') {
    Set-Service -Name $ServiceName -StartupType Manual
}

# Revert PDTWiFi processes
foreach ($process in $PDTWiFiProcesses) {
    $p = Get-Process -Name $process -ErrorAction SilentlyContinue
    if (!$p) {
        Start-Process -FilePath "C:\Program Files (x86)\StationMaster\$process.exe"
    }
}

Write-Host "All tasks completed successfully." -ForegroundColor Green

# Calculate and display script execution time
$endTime = Get-Date
$executionTime = $endTime - $startTime
$totalMinutes = [math]::Floor($executionTime.TotalMinutes)
$totalSeconds = $executionTime.Seconds
Write-Host "Script completed in $($totalMinutes)m $($totalSeconds)s." -ForegroundColor Green
