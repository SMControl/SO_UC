# soua.ps1
# ---
# This script assists in installing Smart Office.
# It performs various checks, downloads necessary files if needed, manages processes and services,
# and sets permissions for specific folders.
# ---
# Version 1.25
# - Improved comments and readability.
# - Added messages for PDTWiFi.exe and PDTWiFi64.exe operations.
# - Added summary at the end of the script.

# Initialize start time
$startTime = Get-Date

# Part 1 - Check for Administrative Privileges
# -----
Write-Host "[Part 1/11] Checking for administrative privileges..." -ForegroundColor Cyan

Function Check-Admin {
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    If (-Not $IsAdmin) {
        Write-Host "ERROR: Please run this script as an administrator." -ForegroundColor Red
        Read-Host "Press any key to exit..."
        Exit
    }
}

Check-Admin

# Part 2 - Check for Running Processes
# -----
Write-Host "[Part 2/11] Checking for running processes..." -ForegroundColor Cyan

Function Check-Process {
    param (
        [string]$ProcessName
    )
    If (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) {
        Write-Host "ERROR: $ProcessName is running. Please close it and restart the script." -ForegroundColor Red
        Read-Host "Press any key to exit..."
        Exit
    }
}

$ProcessesToClose = @("Sm32Main", "Sm32", "PDTWiFi", "PDTWiFi64")
$ProcessesClosed = @()

ForEach ($ProcessName in $ProcessesToClose) {
    If (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) {
        Write-Host "Stopping $ProcessName..." -ForegroundColor Yellow
        Stop-Process -Name $ProcessName -ErrorAction SilentlyContinue
        $ProcessesClosed += $ProcessName
        Write-Host "$ProcessName stopped." -ForegroundColor Green
    }
}

# Part 3 - Create Directory and Change Directory
# -----
Write-Host "[Part 3/11] Creating and changing directory..." -ForegroundColor Cyan

If (-Not (Test-Path -Path "C:\winsm")) {
    New-Item -Path "C:\winsm" -ItemType Directory
}
Set-Location -Path "C:\winsm"

# Part 4 - Download SO_UC.exe if Not Exists
# -----
Write-Host "[Part 4/11] Checking for SO_UC.exe..." -ForegroundColor Cyan

$SO_UC_Path = "C:\winsm\SO_UC.exe"

If (-Not (Test-Path -Path $SO_UC_Path)) {
    # Function to download a file
    Function Download-File {
        param (
            [string]$url,
            [string]$output
        )
        Try {
            Invoke-WebRequest -Uri $url -OutFile $output
        }
        Catch {
            Write-Host "ERROR: Failed to download $url" -ForegroundColor Red
            Read-Host "Press any key to exit..."
            Exit
        }
    }

    $SO_UC_Url = "https://github.com/SMControl/SO_UC/raw/main/SO_UC.exe"
    Write-Host "Downloading SO_UC.exe..." -ForegroundColor Yellow
    Download-File -url $SO_UC_Url -output $SO_UC_Path

    # Ensure the download succeeded
    If (Test-Path -Path $SO_UC_Path) {
        Write-Host "SO_UC.exe downloaded successfully." -ForegroundColor Green
    } Else {
        Write-Host "ERROR: Failed to download SO_UC.exe" -ForegroundColor Red
        Read-Host "Press any key to exit..."
        Exit
    }
} Else {
    Write-Host "SO_UC.exe already exists. Skipping download." -ForegroundColor Green
}

# Part 5 - Run SO_UC.exe and Wait for Completion
# -----
Write-Host "[Part 5/11] Running SO_UC.exe..." -ForegroundColor Cyan

$process = Start-Process -FilePath $SO_UC_Path -PassThru
$process.WaitForExit()
Write-Host "SO_UC.exe has completed." -ForegroundColor Green

# Part 6 - Install Firebird if Necessary
# -----
Write-Host "[Part 6/11] Checking and installing Firebird if necessary..." -ForegroundColor Cyan

If (-Not (Test-Path -Path "C:\Program Files (x86)\Firebird")) {
    Write-Host "Installing Firebird..." -ForegroundColor Yellow
    Invoke-Expression -Command "irm https://raw.githubusercontent.com/SMControl/SM_Firebird_Installer/main/SMFI_Online.ps1 | iex"
}

# Part 7 - Manage Processes and Services
# -----
Write-Host "[Part 7/11] Managing processes and services..." -ForegroundColor Cyan

# Function to manage a service
Function Manage-Service {
    param (
        [string]$ServiceName,
        [string]$Action
    )
    $Service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    If ($Service -ne $null) {
        Switch ($Action) {
            "Stop" {
                If ($Service.Status -eq 'Running') {
                    Write-Host "Stopping $ServiceName service..." -ForegroundColor Yellow
                    Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
                    Write-Host "Waiting for $ServiceName service to stop..." -ForegroundColor Yellow
                    While ((Get-Service -Name $ServiceName).Status -eq 'Stopping') {
                        Start-Sleep -Seconds 1
                    }
                    Write-Host "$ServiceName service stopped." -ForegroundColor Green
                } else {
                    Write-Host "$ServiceName service is already stopped." -ForegroundColor Green
                }
            }
            "Disable" {
                If ($Service.StartType -ne 'Disabled') {
                    Write-Host "Disabling $ServiceName service..." -ForegroundColor Yellow
                    Set-Service -Name $ServiceName -StartupType Disabled
                } else {
                    Write-Host "$ServiceName service is already disabled." -ForegroundColor Green
                }
            }
            "Enable" {
                If ($Service.StartType -eq 'Disabled') {
                    Write-Host "Enabling $ServiceName service..." -ForegroundColor Yellow
                    Set-Service -Name $ServiceName -StartupType Automatic
                } else {
                    Write-Host "$ServiceName service is already enabled." -ForegroundColor Green
                }
            }
            "Start" {
                If ($Service.Status -eq 'Stopped') {
                    Write-Host "Starting $ServiceName service..." -ForegroundColor Yellow
                    Start-Service -Name $ServiceName
                    Write-Host "Waiting for $ServiceName service to start..." -ForegroundColor Yellow
                    While ((Get-Service -Name $ServiceName).Status -eq 'Starting') {
                        Start-Sleep -Seconds 1
                    }
                    Write-Host "$ServiceName service started successfully." -ForegroundColor Green
                } else {
                    Write-Host "$ServiceName service is already started." -ForegroundColor Green
                }
            }
        }
    } else {
        Write-Host "$ServiceName service does not exist. Ignoring." -ForegroundColor Yellow
    }
}

$ServiceName = "srvSOLiveSales"

# Check if srvSOLiveSales service is running before managing
$Service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
If ($Service -ne $null -and $Service.Status -eq 'Running') {
    # Manage srvSOLiveSales service
    Manage-Service -ServiceName $ServiceName -Action "Stop"
    Manage-Service -ServiceName $ServiceName -Action "Disable"
} Else {
    Write-Host "$ServiceName service is not running. Skipping management." -ForegroundColor Green
}

# Ensure only one instance of firebird.exe is running
Do {
    $firebirdInstances = Get-Process -Name "firebird" -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count
    If ($firebirdInstances -gt 1) {
        Write-Host "There are $firebirdInstances instances of firebird.exe running. Please resolve this." -ForegroundColor Red
        Read-Host "Press any key after resolving..."
    }
} Until ($firebirdInstances -le 1)

Stop-Process -Name "firebird" -Force -ErrorAction SilentlyContinue

# Part 8 - Run Setup Executable
# -----
Write-Host "[Part 8/11] Running setup executable..." -ForegroundColor Cyan

# Launch setup executable
$SetupExecutable = Get-ChildItem -Path "C:\winsm\SmartOffice_Installer" -Filter "Setup*.exe" | Select-Object -First 1
Start-Process -FilePath $SetupExecutable.FullName -Wait

# Part 9 - Ensure Smart Office Processes are Closed
# -----
Write-Host "[Part 9/11] Ensuring Smart Office processes are closed..." -ForegroundColor Cyan

Do {
    Check-Process -ProcessName "Sm32"
    Check-Process -ProcessName "Sm32Main"
} Until (-Not (Get-Process -Name "Sm32" -ErrorAction SilentlyContinue) -And -Not (Get-Process -Name "Sm32Main" -ErrorAction SilentlyContinue))

# Part 10 - Set Permissions for StationMaster Folder
# -----
Write-Host "[Part 10/11] Setting permissions for StationMaster folder..." -ForegroundColor Cyan

# Suppress output

 of changing StationMaster folder permissions
icacls "C:\Program Files (x86)\StationMaster" /grant "*S-1-1-0:(OI)(CI)F" /T /C | Out-Null

# Part 11 - Revert Changes
# -----
Write-Host "[Part 11/11] Reverting changes..." -ForegroundColor Cyan

# Restart srvSOLiveSales service if it was running
If ($Service -ne $null -and $Service.Status -eq 'Running') {
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
