# SOUA.ps1
# ---
# This script assists in installing Smart Office.
# It ensures necessary prerequisites are met, processes are managed, and services are configured.
# ---
# Version 1.37
# - Simplified part messages to single lines without additional [OK] messages

# Initialize script start time
$startTime = Get-Date

# Part 1 - Check for Admin Rights
# -----
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Admin)) {
    pause
    exit
} else {
    Write-Host "[Part 1/11] Checking for admin rights..."
}

# Part 2 - Check for Running Smart Office Processes
# -----
$processesToCheck = @("Sm32Main", "Sm32")
foreach ($process in $processesToCheck) {
    if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
        pause
        exit
    }
}
Write-Host "[Part 2/11] Checking for running Smart Office processes..."

# Part 3 - Create Directory if it Doesn't Exist
# -----
$workingDir = "C:\winsm"
if (-not (Test-Path $workingDir)) {
    New-Item -Path $workingDir -ItemType Directory | Out-Null
}
Write-Host "[Part 3/11] Ensuring working directory exists..."

# Part 4 - Download and Run SO_UC.exe
# -----
$SO_UC_Path = "$workingDir\SO_UC.exe"
$SO_UC_URL = "https://github.com/SMControl/SO_UC/blob/main/SO_UC.exe"
if (-not (Test-Path $SO_UC_Path)) {
    Invoke-WebRequest -Uri $SO_UC_URL -OutFile $SO_UC_Path
}
Start-Process -FilePath $SO_UC_Path -Wait
Write-Host "[Part 4/11] Downloading and running SO_UC.exe if necessary..."

# Part 5 - Check for Firebird Installation
# -----
$firebirdDir = "C:\Program Files (x86)\Firebird"
$firebirdInstallerURL = "https://raw.githubusercontent.com/SMControl/SM_Firebird_Installer/main/SMFI_Online.ps1"
if (-not (Test-Path $firebirdDir)) {
    Invoke-Expression -Command (irm $firebirdInstallerURL | iex)
}
Write-Host "[Part 5/11] Checking for Firebird installation..."

# Part 5.1 - Stop SMUpdates.exe if running
# -----
Stop-Process -Name "SMUpdates" -ErrorAction SilentlyContinue
Write-Host "[Part 5.1/11] Checking and stopping SMUpdates.exe if running..."

# Part 6 - Check and Manage Smart Office Live Sales Service
# -----
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
Write-Host "[Part 6/11] Checking and managing Smart Office Live Sales service..."

# Part 7 - Check and Manage PDTWiFi Processes
# -----
$PDTWiFiProcesses = @("PDTWiFi", "PDTWiFi64")
foreach ($process in $PDTWiFiProcesses) {
    $p = Get-Process -Name $process -ErrorAction SilentlyContinue
    if ($p) {
        Stop-Process -Name $process -Force -ErrorAction SilentlyContinue
    }
}
Write-Host "[Part 7/11] Checking and managing PDTWiFi processes..."

# Part 8 - Launch Setup Executable
# -----
$setupDir = "$workingDir\SmartOffice_Installer"
$setupExe = Get-ChildItem -Path $setupDir -Filter "Setup*.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if ($setupExe) {
    Start-Process -FilePath $setupExe.FullName -Wait
}
Write-Host "[Part 8/11] Launching Smart Office setup executable..."

# Part 9 - Wait for User Confirmation
# -----
Read-Host "When Smart Office Installation is fully complete, press Enter to finish off Installation Assistant tasks..."

# Part 10 - Set Permissions for StationMaster Folder
# -----
Stop-Process -Name "SMUpdates" -ErrorAction SilentlyContinue
& icacls "C:\Program Files (x86)\StationMaster" /grant "*S-1-1-0:(OI)(CI)F" /T /C > $null
Write-Host "[Part 10/11] Setting permissions for StationMaster folder..."

# Part 11 - Revert Services and Processes to Original State
# -----
if ($initialServiceState -eq 'Running') {
    Set-Service -Name $ServiceName -StartupType Automatic
    Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
}

foreach ($process in $PDTWiFiProcesses) {
    if ($initialPDTWiFiStates[$process]) {
        Start-Process -FilePath "C:\Program Files (x86)\StationMaster\$process.exe"
    }
}
Write-Host "[Part 11/11] Reverting services and processes to original state..."

# Calculate and display script execution time
$endTime = Get-Date
$executionTime = $endTime - $startTime
$totalMinutes = [math]::Floor($executionTime.TotalMinutes)
$totalSeconds = $executionTime.Seconds
Write-Host "Script completed in $($totalMinutes)m $($totalSeconds)s." -ForegroundColor Yellow
