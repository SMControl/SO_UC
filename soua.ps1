# soua.ps1
# ---
# This script assists in installing Smart Office.
# It performs various checks, downloads necessary files if needed, and manages processes.
# ---
# Version 1.12
# - Enhanced handling of srvSOLiveSales service according to requirements.

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
        Stop-Process -Name $ProcessName -ErrorAction SilentlyContinue
        $ProcessesClosed += $ProcessName
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

    $SO_UC_Url = "https://github.com/SMControl/SO_UC/blob/main/SO_UC.exe"
    Write-Host "Downloading SO_UC.exe..." -ForegroundColor Yellow
    $downloadJob = Start-Job -ScriptBlock { Download-File -url $using:SO_UC_Url -output $using:SO_UC_Path }

    # Wait for the download job to complete
    $downloadJob | Wait-Job

    # Ensure the download succeeded
    If ((Get-Job -Id $downloadJob.Id).State -eq 'Completed') {
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

# Function to stop and disable the service if it exists and is running
Function Manage-Service {
    param (
        [string]$ServiceName
    )
    $Service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    If ($Service -ne $null -and $Service.Status -eq 'Running') {
        Stop-Service -Name $ServiceName
        Set-Service -Name $ServiceName -StartupType Disabled
        Write-Host "Stopped and disabled $ServiceName service." -ForegroundColor Yellow
        return $true
    }
    return $false
}

$ServiceName = "srvSOLiveSales"
$ServiceManaged = $false

# Check and manage srvSOLiveSales service
If (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue) {
    $ServiceManaged = Manage-Service -ServiceName $ServiceName
} Else {
    Write-Host "$ServiceName service does not exist. Ignoring." -ForegroundColor Yellow
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

icacls "C:\Program Files (x86)\StationMaster" /grant "*S-1-1-0:(OI)(CI)F" /T /C > $null 2>&1

# Part 11 - Revert Services to Original State
# -----
Write-Host "[Part 11/11] Reverting services to original state..." -ForegroundColor Cyan

# Function to start and set the service back to its original state
Function Revert-Service {
    param (
        [string]$ServiceName,
        [bool]$WasRunning
    )
    $Service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    If ($Service -ne $null) {
        If ($WasRunning) {
            Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
            Write-Host "Started $ServiceName service." -ForegroundColor Yellow
        }
        Set-Service -Name $ServiceName -StartupType Automatic
        Write-Host "Set $ServiceName service to Automatic startup." -ForegroundColor Yellow
    }
}

# Revert srvSOLiveSales service if managed
If ($ServiceManaged) {
    Revert-Service -ServiceName $ServiceName -WasRunning $true
}

# Initialize end time
$endTime = Get-Date

# Calculate and display total script run time
$totalTime = New-TimeSpan -Start $startTime -End $endTime
Write-Host "Script completed in $($totalTime.ToString('hh\:mm\:ss'))" -ForegroundColor Green

# Summary
Write-Host "`nSummary:" -ForegroundColor Yellow
Write-Host "Processes Closed: $($ProcessesClosed -join ', ')" -ForegroundColor Yellow
If ($ServiceManaged) {
    Write-Host "Service Status ($

ServiceName): $((Get-Service -Name $ServiceName -ErrorAction SilentlyContinue).Status)" -ForegroundColor Yellow
} Else {
    Write-Host "Service $ServiceName was not managed." -ForegroundColor Yellow
}
Read-Host "Press any key to exit..."
