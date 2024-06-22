# SOUA.ps1
# ---
# This script assists in installing Smart Office.
# It ensures necessary prerequisites are met, processes are managed, and services are configured.
# ---
# Version 1.54
# - Added deletion of scheduled task and flag file at script end.
# - Hide SO_UC.exe
# - Ensure scheduled task runs from C:\winsm
# - message to ask to allow SO_UC.exe through firewall to get the latest installer.

# Initialize script start time
$startTime = Get-Date

# Define the flag file path
$flagFilePath = "C:\winsm\SOUA_Flag.txt"
$taskName = "SmartOfficeInstallerResume"
$taskAction = "PowerShell.exe"
$taskArguments = "-ExecutionPolicy Bypass -Command `"irm https://raw.githubusercontent.com/SMControl/SO_UC/main/soua.ps1 | iex`""

# Function to update flag file
function Update-FlagFile {
    param (
        [string]$step,
        [string]$serviceState,
        [string]$processesStopped
    )
    $flagData = @{
        Step = $step
        ServiceState = $serviceState
        ProcessesStopped = $processesStopped
    }
    $flagData | ConvertTo-Json | Set-Content -Path $flagFilePath -Force
}

# Function to read flag file
function Read-FlagFile {
    if (Test-Path $flagFilePath) {
        return Get-Content -Path $flagFilePath | ConvertFrom-Json
    } else {
        return @{
            Step = 0
            ServiceState = @{}
            ProcessesStopped = @()
        }
    }
}

# Function to create the scheduled task
function Create-ScheduledTask {
    $action = New-ScheduledTaskAction -Execute $taskAction -Argument $taskArguments -WorkingDirectory "C:\winsm"
    $trigger = New-ScheduledTaskTrigger -AtStartup
    $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
    Register-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -TaskName $taskName -Description "Resume Smart Office installation script at startup" -Force
}

# Function to delete the scheduled task
function Delete-ScheduledTask {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

# Create the scheduled task (only once at the start)
if (!(Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue)) {
    Create-ScheduledTask
}

# Determine the step to start from
$startStep = Read-FlagFile | Select-Object -ExpandProperty Step

# Part 1 - Check for Admin Rights
# -----
if ($startStep -le 1) {
    Write-Host "[Part 1/12] Checking for admin rights..." -ForegroundColor Green
    function Test-Admin {
        $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }

    if (-not (Test-Admin)) {
        pause
        exit
    }

    Update-FlagFile -step 2 -serviceState @{} -processesStopped @()
}

# Part 2 - Check for Running Smart Office Processes
# -----
if ($startStep -le 2) {
    Write-Host "[Part 2/12] Checking for running Smart Office processes..." -ForegroundColor Green
    $processesToCheck = @("Sm32Main", "Sm32")
    foreach ($process in $processesToCheck) {
        if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
            pause
            exit
        }
    }

    Update-FlagFile -step 3 -serviceState @{} -processesStopped @()
}

# Part 3 - Create Directory if it Doesn't Exist
# -----
if ($startStep -le 3) {
    Write-Host "[Part 3/12] Ensuring working directory exists..." -ForegroundColor Green
    $workingDir = "C:\winsm"
    if (-not (Test-Path $workingDir)) {
        New-Item -Path $workingDir -ItemType Directory | Out-Null
    }

    Update-FlagFile -step 4 -serviceState @{} -processesStopped @()
}

# Part 4 - Download and Run SO_UC.exe Hidden if Necessary
# -----
Write-Host "[Part 4/12] Downloading latest Smart Office Setup if necessary..." -ForegroundColor Green

# Display message about firewall
Write-Host "You may need to allow SO_UC.exe through the Firewall. Its job is to get the latest version of Smart Office Setup." -ForegroundColor Cyan

$SO_UC_Path = "$workingDir\SO_UC.exe"
$SO_UC_URL = "https://github.com/SMControl/SO_UC/raw/main/SO_UC.exe"
if (-not (Test-Path $SO_UC_Path)) {
    Invoke-WebRequest -Uri $SO_UC_URL -OutFile $SO_UC_Path
}

# Start SO_UC.exe hidden
$processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
$processStartInfo.FileName = $SO_UC_Path
$processStartInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

Start-Process -FilePath $processStartInfo.FileName -Wait -WindowStyle $processStartInfo.WindowStyle

# Part 5 - Check for Firebird Installation
# -----
if ($startStep -le 5) {
    Write-Host "[Part 5/12] Checking for Firebird installation..." -ForegroundColor Green
    $firebirdDir = "C:\Program Files (x86)\Firebird"
    $firebirdInstallerURL = "https://raw.githubusercontent.com/SMControl/SM_Firebird_Installer/main/SMFI_Online.ps1"
    if (-not (Test-Path $firebirdDir)) {
        Invoke-Expression -Command (irm $firebirdInstallerURL | iex)
    }

    Update-FlagFile -step 6 -serviceState @{} -processesStopped @()
}

# Part 6 - Stop SMUpdates.exe if Running
# -----
if ($startStep -le 6) {
    Write-Host "[Part 6/12] Checking and stopping SMUpdates.exe if running..." -ForegroundColor Green
    Stop-Process -Name "SMUpdates" -ErrorAction SilentlyContinue

    Update-FlagFile -step 7 -serviceState @{} -processesStopped @("SMUpdates")
}

# Part 7 - Check and Manage Smart Office Live Sales Service
# -----
if ($startStep -le 7) {
    Write-Host "[Part 7/12] Checking and managing Smart Office Live Sales service..." -ForegroundColor Green
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

    $serviceState = @{ $ServiceName = $initialServiceState }
    Update-FlagFile -step 8 -serviceState $serviceState -processesStopped @("SMUpdates")
}

# Part 8 - Check and Manage PDTWiFi Processes
# -----
if ($startStep -le 8) {
    Write-Host "[Part 8/12] Checking and managing PDTWiFi processes..." -ForegroundColor Green
    $PDTWiFiProcesses = @("PDTWiFi", "PDTWiFi64")
    foreach ($process in $PDTWiFiProcesses) {
        $p = Get-Process -Name $process -ErrorAction SilentlyContinue
        if ($p) {
            Stop-Process -Name $process -Force -ErrorAction SilentlyContinue
        }
    }

    Update-FlagFile -step 9 -serviceState $serviceState -processesStopped $PDTWiFiProcesses
}

# Part 9 - Launch Setup Executable
# -----
if ($startStep -le 9) {
    Write-Host "[Part 9/12] Launching Smart Office setup executable..." -ForegroundColor Green
    $setupDir = "$workingDir\SmartOffice_Installer"
    $setupExe = Get-ChildItem -Path $setupDir -Filter "Setup*.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if ($setupExe) {
        Start-Process -FilePath $setupExe.FullName -Wait
    }

    Update-FlagFile -step 10 -serviceState $serviceState -processesStopped $PDTWiFiProcesses
}

# Part 10 - Wait for User Confirmation
# -----
if ($startStep -le 10) {
    Write-Host "[Part 10/12] Please press Enter when the Smart Office installation is FULLY finished..." -ForegroundColor White
    Read-Host

    # Check for Running Smart Office Processes Again
    $processesToCheck = @("Sm32Main", "Sm32")
    $processRunning = $false
    foreach ($process in $processesToCheck) {
        if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
            $processRunning = $true
            Write-Host "Smart Office is still running. Please close it and press Enter to continue..." -ForegroundColor White
            Read-Host
        }
    }

    Update

-FlagFile -step 11 -serviceState $serviceState -processesStopped $PDTWiFiProcesses
}

# Part 11 - Set Permissions for StationMaster Folder
# -----
Write-Host "[Part 11/12] Setting permissions for StationMaster folder..." -ForegroundColor Green
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

# Delete the scheduled task and flag file
Delete-ScheduledTask
Remove-Item -Path $flagFilePath -Force

# Calculate and display script execution time
$endTime = Get-Date
$executionTime = $endTime - $startTime
$totalMinutes = [math]::Floor($executionTime.TotalMinutes)
$totalSeconds = $executionTime.Seconds
Write-Host "Script completed in $($totalMinutes)m $($totalSeconds)s." -ForegroundColor Green
