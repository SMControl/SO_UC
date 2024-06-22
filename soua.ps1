Write-Host "SOUA.ps1 - Version 1.108" -ForegroundColor Green

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

# Part 3 - Download SO_UC.exe if Necessary
# -----
Write-Host "[Part 3] Downloading SO_UC.exe if necessary..." -ForegroundColor Green
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

# Part 4 - Check for Firebird Installation
# -----
Write-Host "[Part 4] Checking for Firebird installation..." -ForegroundColor Green
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

# Part 5 - Stop SMUpdates if Running
# -----
Write-Host "[Part 5] Stopping SMUpdates if running..." -ForegroundColor Green
try {
    Stop-Process -Name "SMUpdates" -ErrorAction SilentlyContinue
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

if ($service -and $service.Status -eq 'Running') {
    $wasRunning = $true
    try {
        Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
        Set-Service -Name $ServiceName -StartupType Disabled
    } catch {
        Write-Host "Error stopping service '$ServiceName': $_" -ForegroundColor Red
        exit
    }
}

# Part 7 - Manage PDTWiFi Processes
# -----
Write-Host "[Part 7] Managing PDTWiFi processes..." -ForegroundColor Green
$PDTWiFiProcesses = @("PDTWiFi", "PDTWiFi64")
$PDTWiFiStates = @{}

foreach ($process in $PDTWiFiProcesses) {
    $p = Get-Process -Name $process -ErrorAction SilentlyContinue
    if ($p) {
        $PDTWiFiStates[$process] = $p.Status
        Stop-Process -Name $process -Force -ErrorAction SilentlyContinue
    }
}

# Part 8 - Launch Setup Executable
# -----
Write-Host "[Part 8] Launching Smart Office setup executable..." -ForegroundColor Green
$setupDir = "$workingDir\SmartOffice_Installer"
if (-not (Test-Path $setupDir -PathType Container)) {
    Write-Host "Error: Setup directory '$setupDir' does not exist." -ForegroundColor Red
    exit
}

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

# Part 9 - Wait for User Confirmation
# -----
Write-Host "[Part 9] When Upgrade is FULLY complete, Press Enter...." -ForegroundColor Yellow
Read-Host

# Check for Running Smart Office Processes Again
$processesToCheck = @("Sm32Main", "Sm32")
foreach ($process in $processesToCheck) {
    if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
        Write-Host "Smart Office is still running. Please close it and press Enter to continue..." -ForegroundColor Red
        Read-Host
    }
}

# Kill PDTWiFi and SMUpdates if running (after a reboot for example)
$processesToKill = @("PDTWiFi", "PDTWiFi64", "SMUpdates")
foreach ($process in $processesToKill) {
    Stop-Process -Name $process -Force -ErrorAction SilentlyContinue
}

# Part 10 - Set Permissions for StationMaster Folder
# -----
Write-Host "[Part 10] Setting permissions for StationMaster folder..." -ForegroundColor Green
try {
    & icacls "C:\Program Files (x86)\StationMaster" /grant "*S-1-1-0:(OI)(CI)F" /T /C > $null
} catch {
    Write-Host "Error setting permissions for StationMaster folder: $_" -ForegroundColor Red
}

# Part 11 - Set Permissions for Firebird Folder
# -----
Write-Host "[Part 11] Setting permissions for Firebird folder..." -ForegroundColor Green
try {
    & icacls "C:\Program Files (x86)\Firebird" /grant "*S-1-1-0:(OI)(CI)F" /T /C > $null
} catch {
    Write-Host "Error setting permissions for Firebird folder: $_" -ForegroundColor Red
}

# Part 12 - Revert Smart Office Live Sales Service
# -----
Write-Host "[Part 12] Reverting Smart Office Live Sales service..." -ForegroundColor Green
if ($wasRunning) {
    try {
        Set-Service -Name $ServiceName -StartupType Automatic
        Start-Service -Name $ServiceName
    } catch {
        Write-Host "Error reverting service '$ServiceName' to its previous state: $_" -ForegroundColor Red
    }
}

# Part 13 - Revert PDTWiFi Processes
# -----
Write-Host "[Part 13] Reverting PDTWiFi processes..." -ForegroundColor Green
foreach ($process in $PDTWiFiProcesses) {
    if ($PDTWiFiStates.ContainsKey($process)) {
        $status = $PDTWiFiStates[$process]
        if ($status -eq 'Running') {
            try {
                Start-Process -FilePath "C:\Path\to\$process.exe" -ErrorAction SilentlyContinue
            } catch {
                Write-Host "Error starting $process $_" -ForegroundColor Red
            }
        }
    }
}

# Part 14 - Clean Up and Finish
# -----
Write-Host "[Part 14] Cleaning up and finishing script..." -ForegroundColor Green

# Calculate and display script execution time
$endTime = Get-Date
$executionTime = $endTime - $startTime
$totalMinutes = [math]::Floor($executionTime.TotalMinutes)
$totalSeconds = $executionTime.Seconds
Write-Host "Script completed successfully in $($totalMinutes)m $($totalSeconds)s." -ForegroundColor Green
