Write-Host "SOUA.ps1" -ForegroundColor Green
# ---
# This script assists in installing Smart Office.
# It ensures necessary prerequisites are met, processes are managed, and services are configured.
# ---
Write-Host "Version 1.84" -ForegroundColor Green
# - Added logging of upgrade timing information
# - Added display of upgrade statistics at script start

# Initialize script start time
$startTime = Get-Date

# Initialize log directory
$logDir = "C:\winsm\SmartOffice_Installer\Update_Assistant_Logs_and_Records"
if (-not (Test-Path $logDir -PathType Container)) {
    try {
        New-Item -Path $logDir -ItemType Directory -ErrorAction Stop | Out-Null
    } catch {
        Write-Host "Error creating log directory: $_" -ForegroundColor Red
        exit
    }
}

# Display upgrade statistics
$logFilePath = "$logDir\Upgrade_Log.txt"
if (Test-Path $logFilePath) {
    $logEntries = Get-Content -Path $logFilePath
    $totalUpgrades = $logEntries.Count

    if ($totalUpgrades -gt 0) {
        $durations = $logEntries | ForEach-Object {
            if ($_ -match 'Duration: (\d+)m (\d+)s') {
                $minutes = [int]$matches[1]
                $seconds = [int]$matches[2]
                $minutes * 60 + $seconds
            }
        }

        $shortest = $durations | Measure-Object -Minimum | Select-Object -ExpandProperty Minimum
        $longest = $durations | Measure-Object -Maximum | Select-Object -ExpandProperty Maximum
        $average = [math]::Round(($durations | Measure-Object -Average).Average)
        $mean = [math]::Round(($durations | Measure-Object -Median).Median)

        Write-Host "Upgrade Statistics:" -ForegroundColor Green
        Write-Host "-------------------" -ForegroundColor Green
        Write-Host "Total Number of Upgrades Performed: $totalUpgrades"
        Write-Host "Shortest Upgrade: $(int)$(($shortest - $shortest % 60) / 60)m $(($shortest % 60))s"
        Write-Host "Longest Upgrade: $(int)$(($longest - $longest % 60) / 60)m $(($longest % 60))s"
        Write-Host "Average Upgrade: $(int)$(($average - $average % 60) / 60)m $(($average % 60))s"
        Write-Host "Mean Upgrade: $(int)$(($mean - $mean % 60) / 60)m $(($mean % 60))s"
        Write-Host "-------------------" -ForegroundColor Green
        Write-Host " "
    }
}

# Define the flag file path
$flagFilePath = "$winsmDir\SOUA_Flag.txt"

# Check if flag file exists to determine starting point
if (Test-Path $flagFilePath -PathType Leaf) {
    $startStep = 10
    Write-Host "Resuming from previous position..." -ForegroundColor Yellow
} else {
    $startStep = 0
    # Create flag file silently
    try {
        New-Item -Path $flagFilePath -ItemType File -ErrorAction Stop | Out-Null
    } catch {
        # Handle flag file creation error silently
        exit
    }
}

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
if ($startStep -le 2) {
    Write-Host "[Part 2/13] Checking for running Smart Office processes..." -ForegroundColor Green
    $processesToCheck = @("Sm32Main", "Sm32")
    foreach ($process in $processesToCheck) {
        if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
            Write-Host "Error: Smart Office process '$process' is running. Close it and retry." -ForegroundColor Red
            pause
            exit
        }
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
if ($startStep -le 5) {
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
}

# Part 6 - Stop SMUpdates.exe if Running
# -----
if ($startStep -le 6) {
    Write-Host "[Part 6/13] Checking and stopping SMUpdates.exe if running..." -ForegroundColor Green
    try {
        Stop-Process -Name "SMUpdates" -ErrorAction SilentlyContinue
    } catch {
        Write-Host "Error stopping SMUpdates.exe: $_" -ForegroundColor Red
        exit
    }
}

# Part 7 - Check and Manage Smart Office Live Sales Service
# -----
if ($startStep -le 7) {
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
}

# Part 8 - Check and Manage PDTWiFi Processes
# -----
if ($startStep -le 8) {
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
}

# Part 9 - Launch Setup Executable
# -----
if ($startStep -le 9) {
    Write-Host "[Part 9/13] Launching Smart Office setup executable..." -ForegroundColor Green
    $setupDir = "$workingDir\SmartOffice_Installer"
    $setupExe = Get-ChildItem -Path $setupDir -Filter "Setup*.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if ($setupExe) {
        Start-Process -FilePath $setupExe.FullName -Wait
    }
}

# Part 10 - Wait for User Confirmation
# -----
Write-Host "[Part 10/13] Post Installation" -ForegroundColor Green
if ($startStep -le 10) {
    Write-Host "[Part 10/13]

 Please press Enter when the Smart Office installation is FULLY finished..." -ForegroundColor White
    Read-Host

    # Check for Running Smart Office Processes Again
    $processesToCheck = @("Sm32Main", "Sm32")
    foreach ($process in $processesToCheck) {
        if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
            Write-Host "Smart Office is still running. Please close it and press Enter to continue..." -ForegroundColor Red
            Read-Host
        }
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

# Log timestamp, duration, and setup file used
$logEntry = "Timestamp: $(Get-Date)"
$logEntry += "`nDuration: $($totalMinutes)m $($totalSeconds)s"
$logEntry += "`nSetup File: $($setupExe.Name)"
Add-Content -Path $logFilePath -Value $logEntry

# Delete the flag file
Remove-Item -Path $flagFilePath -Force -ErrorAction SilentlyContinue

Write-Host " "

# Calculate and display script execution time
$endTime = Get-Date
$executionTime = $endTime - $startTime
$totalMinutes = [math]::Floor($executionTime.TotalMinutes)
$totalSeconds = $executionTime.Seconds
Write-Host "Script completed successfully in $($totalMinutes)m $($totalSeconds)s." -ForegroundColor Green
