Write-Host "SOUA.ps1 - Version 1.123" -ForegroundColor Green
# ---
# - removed un-nessecary messages
# - colour change
# - typos
# - spacing

# Initialize script start time
$startTime = Get-Date

# Set the working directory - move to after admin check
$workingDir = "C:\winsm"
if (-not (Test-Path $workingDir -PathType Container)) {
    try {
        New-Item -Path $workingDir -ItemType Directory -ErrorAction Stop | Out-Null
    } catch {
        Write-Host "Error: Unable to create directory $workingDir" -ForegroundColor Red
        exit
    }
}

Set-Location -Path $workingDir

Write-Host "[WARNING] Upgrades requiring a reboot are not yet supported." -ForegroundColor Red
Write-Host "[WARNING] If a Reboot is required, Post Upgrade Tasks must be performed manually." -ForegroundColor Yellow

# Part 1 - Check for Admin Rights
# -----
Write-Host "[Part 1/15] System Pre-Checks" -ForegroundColor Cyan
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Admin)) {
    Write-Host "Error: Administrator rights required to run this script. Exiting." -ForegroundColor Red
    pause
    exit
}

# Part 2 - Check for Running SO Processes
# -----
Write-Host "[Part 2/15] Checking processes" -ForegroundColor Cyan
$processesToCheck = @("Sm32Main", "Sm32")
foreach ($process in $processesToCheck) {
    if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
        Write-Host "Error: SO process '$process' is running. Close it and retry." -ForegroundColor Red
        pause
        exit
    }
}

# Part 3 - Download SO_UC.exe if Necessary and Launching SO_UC.exe and Wait for Completion
# Part 3 - PartVersion 1.02
# -----
Write-Host "[Part 3/15] Checking SO_UC" -ForegroundColor Cyan
$SO_UC_Path = "$workingDir\SO_UC.exe"
$SO_UC_URL = "https://github.com/SMControl/SO_UC/raw/main/SO_UC.exe"
if (-not (Test-Path $SO_UC_Path)) {
    Write-Host "Not found. Downloading..." -ForegroundColor Yellow
    try {
        $download = Invoke-WebRequest -Uri $SO_UC_URL -OutFile $SO_UC_Path -PassThru
        Write-Host "Downloading SO_UC..." -ForegroundColor Yellow
        while ($download.IsCompleted -eq $false) {
            Write-Host "Progress: $($download.BytesReceived / $download.TotalBytes * 100)% complete" -ForegroundColor Yellow
            Start-Sleep -Seconds 1
        }
        Write-Host "Downloaded successfully." -ForegroundColor Green
    } catch {
        Write-Host "Error downloading SO_UC: $_" -ForegroundColor Red
        exit
    }
} else {
}

# Part 3.1 - Launching SO_UC.exe hidden and wait for completion
# -----
Write-Host "Launching SO_UC.exe. Please allow through Firewall" -ForegroundColor Green
$process = Start-Process -FilePath $SO_UC_Path -PassThru -WindowStyle Hidden
if ($process) {
    Write-Host "Checking/Downloading latest version of SO Installer from SM. Please wait..." -ForegroundColor Green
    $process.WaitForExit()
    Start-Sleep -Seconds 2
} else {
    Write-Host "Failed to start SO_UC.exe." -ForegroundColor Red
    exit
}

# Part 4 - Check for Firebird Installation
# -----
Write-Host "[Part 4/15] Checking for Firebird installation" -ForegroundColor Cyan
$firebirdDir = "C:\Program Files (x86)\Firebird"
$firebirdInstallerURL = "https://raw.githubusercontent.com/SMControl/SM_Firebird_Installer/main/SMFI_Online.ps1"
if (-not (Test-Path $firebirdDir)) {
    Write-Host "Firebird not found. Installing Firebird..." -ForegroundColor Yellow
    try {
        Invoke-Expression -Command (irm $firebirdInstallerURL | iex)
        Write-Host "Firebird installed successfully." -ForegroundColor Green
    } catch {
        Write-Host "Error installing Firebird: $_" -ForegroundColor Red
        exit
    }
} else {
}

# Part 5 - Stop SMUpdates if Running
# -----
Write-Host "[Part 5/15] Stopping SMUpdates if running" -ForegroundColor Cyan
try {
    $smUpdatesProcess = Get-Process -Name "SMUpdates" -ErrorAction SilentlyContinue
    if ($smUpdatesProcess) {
        Write-Host "Stopping SMUpdates process..." -ForegroundColor Yellow
        Stop-Process -Name "SMUpdates" -Force -ErrorAction SilentlyContinue
        Write-Host "SMUpdates stopped successfully." -ForegroundColor Green
    } else {
    }
} catch {
    Write-Host "Error stopping SMUpdates.exe: $_" -ForegroundColor Red
    exit
}

# Part 6 - Manage SO Live Sales Service
# -----
Write-Host "[Part 6/15] Managing SO Live Sales service" -ForegroundColor Cyan
$ServiceName = "srvSOLiveSales"
$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
$wasRunning = $false

if ($service) {
    if ($service.Status -eq 'Running') {
        $wasRunning = $true
        Write-Host "Stopping $ServiceName service..." -ForegroundColor Yellow
        try {
            Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
            Set-Service -Name $ServiceName -StartupType Disabled
            Write-Host "$ServiceName service stopped successfully and set to Disabled." -ForegroundColor Green
        } catch {
            Write-Host "Error stopping service '$ServiceName': $_" -ForegroundColor Red
            exit
        }
    } else {
        Write-Host "$ServiceName service is not running." -ForegroundColor Yellow
    }
} else {
    Write-Host "$ServiceName service not found." -ForegroundColor Yellow
}

# Part 7 - Manage PDTWiFi Processes and Log State
# -----
Write-Host "[Part 7/15] Managing PDTWiFi processes" -ForegroundColor Cyan
$PDTWiFiProcesses = @("PDTWiFi", "PDTWiFi64")
$PDTWiFiStates = @{}

foreach ($process in $PDTWiFiProcesses) {
    $p = Get-Process -Name $process -ErrorAction SilentlyContinue
    if ($p) {
        $PDTWiFiStates[$process] = $p.Status
        Write-Host "Stopping $process process..." -ForegroundColor Yellow
        Stop-Process -Name $process -Force -ErrorAction SilentlyContinue
        Write-Host "$process stopped successfully." -ForegroundColor Green
    } else {
        Write-Host "$process process is not running." -ForegroundColor Yellow
        # If the process was not running, ensure it's reflected in the states
        $PDTWiFiStates[$process] = "Not running"
    }
}

# Check again and update states for processes that were running before stopping
foreach ($process in $PDTWiFiProcesses) {
    if ($PDTWiFiStates[$process] -eq $null) {
        $PDTWiFiStates[$process] = "Running"
    }
}

# Log PDTWiFi states to a temporary file
$PDTWiFiStatesFilePath = "$workingDir\PDTWiFiStates.txt"
$PDTWiFiStates.GetEnumerator() | ForEach-Object { "$($_.Key): $($_.Value)" } | Out-File -FilePath $PDTWiFiStatesFilePath

# Part 8 - Check and Wait for Single Instance of Firebird.exe
# -----
Write-Host "[Part 8/15] Checking and waiting for a single instance of Firebird" -ForegroundColor Cyan

$setupDir = "$workingDir\SmartOffice_Installer"
if (-not (Test-Path $setupDir -PathType Container)) {
    Write-Host "Error: Setup directory '$setupDir' does not exist." -ForegroundColor Red
    exit
}

function WaitForSingleFirebirdInstance {
    $firebirdProcesses = Get-Process -Name "firebird" -ErrorAction SilentlyContinue
    while ($firebirdProcesses.Count -gt 1) {
        Write-Host "Warning: Multiple instances of 'firebird.exe' are running. Please ensure only one instance is running." -ForegroundColor Yellow
        Write-Host "Currently running instances: $($firebirdProcesses.Count)" -ForegroundColor Yellow
        Start-Sleep -Seconds 1
        $firebirdProcesses = Get-Process -Name "firebird" -ErrorAction SilentlyContinue
    }
}

# Call function to wait for a single instance of Firebird
WaitForSingleFirebirdInstance

# Part 9 - Launch SO Setup Executable
# Part 9 - PartVersion 1.06
# -----
Write-Host "[Part 9/15] Proceeding to launch SO setup executable" -ForegroundColor Cyan

$setupExe = Get-ChildItem -Path "C:\winsm\SmartOffice_Installer" -Filter "*.exe" | Select-Object -First 1
if ($setupExe) {
    Write-Host "Found setup executable: $($setupExe.FullName)" -ForegroundColor Green
    try {
        Start-Process -FilePath $setupExe.FullName -Wait
    } catch {
        Write-Host "Error starting setup executable: $_" -ForegroundColor Red
        exit
    }
} else {
    Write-Host "Error: No executable (.exe) found in 'C:\winsm\SmartOffice_Installer'." -ForegroundColor Red
    exit
}

# Part 10 - Wait for User Confirmation
# -----
Write-Host "[Part 10/15] Post Upgrade" -ForegroundColor Cyan
Write-Host " "
Write-Host "    When upgrade is FULLY complete and SO is closed;" -ForegroundColor Magenta
Write-Host "    Press Enter to finish off Post Install tasks..." -ForegroundColor Magenta
Read-Host

# Check for Running SO Processes Again
$processesToCheck = @("Sm32Main", "Sm32")
foreach ($process in $processesToCheck) {
    if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
        Write-Host "SO is still running. Please close it and press Enter to continue..." -ForegroundColor Red
        Read-Host
    }
}

# Part 11 - Set Permissions for SM Folder
# -----
Write-Host "[Part 11/15] Setting permissions for SM folder" -ForegroundColor Cyan
try {
    & icacls "C:\Program Files (x86)\StationMaster" /grant "*S-1-1-0:(OI)(CI)F" /T /C > $null
} catch {
    Write-Host "Error setting permissions for SM folder: $_" -ForegroundColor Red
}

# Part 12 - Set Permissions for Firebird Folder
# -----
Write-Host "[Part 12/15] Setting permissions for Firebird folder" -ForegroundColor Cyan
try {
    & icacls "C:\Program Files (x86)\Firebird" /grant "*S-1-1-0:(OI)(CI)F" /T /C > $null
} catch {
    Write-Host "Error setting permissions for Firebird folder." -ForegroundColor Red
}

# Part 13 - Revert SO Live Sales Service
# -----
Write-Host "[Part 13/15] Reverting SO Live Sales service" -ForegroundColor Cyan
if ($wasRunning) {
    try {
        Write-Host "Setting $ServiceName service back to Automatic startup..." -ForegroundColor Yellow
        Set-Service -Name $ServiceName -StartupType Automatic
        Start-Service -Name $ServiceName
        Write-Host "$ServiceName service reverted to its previous state." -ForegroundColor Green
    } catch {
        Write-Host "Error reverting service '$ServiceName' to its previous state: $_" -ForegroundColor Red
    }
} else {
    Write-Host "$ServiceName service was not running before, so no action needed." -ForegroundColor Yellow
}

# Part 14 - Revert PDTWiFi Processes
# -----
Write-Host "[Part 14/15] Reverting PDTWiFi processes" -ForegroundColor Cyan

# Retrieve PDTWiFi states from the temporary file
if (Test-Path $PDTWiFiStatesFilePath) {
    $storedStates = Get-Content -Path $PDTWiFiStatesFilePath | ForEach-Object {
        $parts = $_ -split ":"
        $process = $parts[0].Trim()
        $status = $parts[1].Trim()
        [PSCustomObject]@{
            Process = $process
            Status = $status
        }
    }
} else {
    Write-Host "Error: PDTWiFiStates.txt not found. Unable to revert PDTWiFi processes." -ForegroundColor Red
}

foreach ($process in $PDTWiFiProcesses) {
    $currentStatus = $storedStates | Where-Object { $_.Process -eq $process } | Select-Object -ExpandProperty Status
    if ($currentStatus -eq 'Running') {
        Write-Host "Starting $process process..." -ForegroundColor Yellow
        try {
            Start-Process -FilePath "C:\Program Files (x86)\StationMaster\$process" -ErrorAction SilentlyContinue
            Write-Host "$process started successfully." -ForegroundColor Green
        } catch {
            Write-Host "Error starting $process $_" -ForegroundColor Red
        }
    } else {
        Write-Host "$process was not running before, so no action needed." -ForegroundColor Yellow
    }
}

# Clean up and remove temporary file
if (Test-Path $PDTWiFiStatesFilePath) {
    Remove-Item -Path $PDTWiFiStatesFilePath -Force
}

# Part 15 - Clean up and Finish Script
# -----
Write-Host "[Part 15/15] Cleaning up and finishing script" -ForegroundColor Cyan

# Calculate and display script execution time
$endTime = Get-Date
$executionTime = $endTime - $startTime
$totalMinutes = [math]::Floor($executionTime.TotalMinutes)
$totalSeconds = $executionTime.Seconds
Write-Host " "
Write-Host "Script completed successfully in $($totalMinutes)m $($totalSeconds)s." -ForegroundColor Green
Write-Host " "
