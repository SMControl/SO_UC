Write-Host "SOUA.ps1 - Version 1.115" -ForegroundColor Green
# ---
# - cleaned up commenting

# Initialize script start time
$startTime = Get-Date

# Set the working directory
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

# Part 1 - Check for Admin Rights
# -----
Write-Host "[Part 1] Checking for admin rights..." -ForegroundColor Green
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
Write-Host "[Part 2] Checking for running Smart Office processes..." -ForegroundColor Green
$processesToCheck = @("Sm32Main", "Sm32")
foreach ($process in $processesToCheck) {
    if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
        Write-Host "Error: Smart Office process '$process' is running. Close it and retry." -ForegroundColor Red
        pause
        exit
    }
}

# Part 3 - Download SO_UC.exe if Necessary and Run it.
# -----
Write-Host "[Part 3] Downloading SO_UC.exe if necessary..." -ForegroundColor Green
$SO_UC_Path = "$workingDir\SO_UC.exe"
$SO_UC_URL = "https://github.com/SMControl/SO_UC/raw/main/SO_UC.exe"
if (-not (Test-Path $SO_UC_Path)) {
    Write-Host "SO_UC.exe not found. Downloading..." -ForegroundColor Yellow
    try {
        Invoke-WebRequest -Uri $SO_UC_URL -OutFile $SO_UC_Path
        Write-Host "Downloaded successful." -ForegroundColor Green
    } catch {
        Write-Host "Error downloading: $_" -ForegroundColor Red
        exit
    }
} else {
}

# Launching SO_UC.exe
Start-Process -FilePath $SO_UC_Path -NoNewWindow
if ($?) {
    Write-Host "Please allow always SO_UC.exe through Firewall." -ForegroundColor Yellow
    Write-Host "This program obtains latest Smart Office setup." -ForegroundColor Yellow
} else {
    Write-Host "Failed to launch SO_UC.exe." -ForegroundColor Red
}

# Part 4 - Check for Firebird Installation
# -----
Write-Host "[Part 4] Checking for Firebird installation..." -ForegroundColor Green
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
Write-Host "[Part 5] Stopping SMUpdates if running..." -ForegroundColor Green
try {
    $smUpdatesProcess = Get-Process -Name "SMUpdates" -ErrorAction SilentlyContinue
    if ($smUpdatesProcess) {
        Write-Host "Stopping SMUpdates process..." -ForegroundColor Yellow
        Stop-Process -Name "SMUpdates" -Force -ErrorAction SilentlyContinue
        Write-Host "SMUpdates stopped successfully." -ForegroundColor Green
    } else {
        Write-Host "SMUpdates is not running." -ForegroundColor Yellow
    }
} catch {
    Write-Host "Error stopping SMUpdates.exe: $_" -ForegroundColor Red
    exit
}

# Part 6 - Manage Smart Office Live Sales Service
# -----
Write-Host "[Part 6] Managing Smart Office Live Sales service..." -ForegroundColor Green
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
Write-Host "[Part 7] Managing PDTWiFi processes..." -ForegroundColor Green
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

Write-Host "PDTWiFi states logged to: $PDTWiFiStatesFilePath" -ForegroundColor Green


# Part 8 - Check and Wait for Single Instance of Firebird.exe
# -----
Write-Host "[Part 8] Checking and waiting for a single instance of 'firebird.exe'..." -ForegroundColor Green

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

# Part 9 - Launch Smart Office Setup Executable
# -----
Write-Host "[Part 9] Proceeding to launch Smart Office setup executable..." -ForegroundColor Green

$setupExe = Get-ChildItem -Path $setupDir -Filter "Setup*.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($setupExe) {
    Write-Host "Found setup executable: $($setupExe.FullName)" -ForegroundColor Green
    try {
        Start-Process -FilePath $setupExe.FullName -Wait
    } catch {
        Write-Host "Error starting setup executable: $_" -ForegroundColor Red
        exit
    }
} else {
    Write-Host "Error: Smart Office setup executable not found in '$setupDir'." -ForegroundColor Red
    exit
}

# Part 10 - Wait for User Confirmation
# -----
Write-Host "[Part 10] When upgrade is FULLY complete, press Enter...." -ForegroundColor Yellow
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
Write-Host "[Part 11] Setting permissions for StationMaster folder..." -ForegroundColor Green
try {
    & icacls "C:\Program Files (x86)\StationMaster" /grant "*S-1-1-0:(OI)(CI)F" /T /C > $null
} catch {
    Write-Host "Error setting permissions for StationMaster folder: $_" -ForegroundColor Red
}

# Part 12 - Set Permissions for Firebird Folder
# -----
Write-Host "[Part 12] Setting permissions for Firebird folder..." -ForegroundColor Green
try {
    & icacls "C:\Program Files (x86)\Firebird" /grant "*S-1-1-0:(OI)(CI)F" /T /C > $null
} catch {
    Write-Host "Error setting permissions for Firebird folder." -ForegroundColor Red
}

# Part 13 - Revert Smart Office Live Sales Service
# -----
Write-Host "[Part 13] Reverting Smart Office Live Sales service..." -ForegroundColor Green
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
Write-Host "[Part 14] Reverting PDTWiFi processes..." -ForegroundColor Green

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
    Write-Host "Removed temporary file: $PDTWiFiStatesFilePath" -ForegroundColor Green
}


# Part 15 - Clean up and Finish Script
# -----
Write-Host "[Part 15] Cleaning up and finishing script..." -ForegroundColor Green

# Remove temporary files
if (Test-Path $PDTWiFiStatesFilePath) {
    Remove-Item -Path $PDTWiFiStatesFilePath -Force
    Write-Host "Removed temporary file: $PDTWiFiStatesFilePath" -ForegroundColor Green
}

# Calculate and display script execution time
$endTime = Get-Date
$executionTime = $endTime - $startTime
$totalMinutes = [math]::Floor($executionTime.TotalMinutes)
$totalSeconds = $executionTime.Seconds
Write-Host "Script completed successfully in $($totalMinutes)m $($totalSeconds)s." -ForegroundColor Green
