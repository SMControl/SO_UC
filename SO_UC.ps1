Write-Host "SO_UC.ps1 - Version 1.05"
# -----
# - Initial check for scheduled task
# - Fetch and filter .exe links
# - Download and manage the latest Setup.exe
# - Cleanup older downloads

# Part 1 - Check if scheduled task exists and create if it doesn't
# PartVersion 1.00
# -----
# Check if the scheduled task exists
$taskExists = Get-ScheduledTask -TaskName "SO InstallerUpdates" -ErrorAction SilentlyContinue

# If the task does not exist, proceed to create it
if (-not $taskExists) {
    # Define the action to run the program
    $action = New-ScheduledTaskAction -Execute "C:\winsm\SO_UC.exe"

    # Define a random time between 00:00 and 05:59 for daily execution
    $randomHour = Get-Random -Minimum 0 -Maximum 5
    $randomMinute = Get-Random -Minimum 0 -Maximum 59
    $trigger = New-ScheduledTaskTrigger -Daily -At "${randomHour}:${randomMinute}"

    # Define task settings with hidden property set to true
    $settings = New-ScheduledTaskSettingsSet -Hidden:$true

    # Create the scheduled task with elevated privileges and hidden
    Register-ScheduledTask -TaskName "SO InstallerUpdates" -Action $action -Trigger $trigger -Settings $settings -RunLevel Highest
}

# Part 2 - Retrieve .exe links from the webpage
# PartVersion 1.00
# -----
# Retrieve .exe links from the webpage
$exeLinks = (Invoke-WebRequest -Uri "https://www.stationmaster.com/downloads/").Links | Where-Object { $_.href -match "\.exe$" } | ForEach-Object { $_.href }

# Part 3 - Filter for the highest version of Setup.exe
# PartVersion 1.00
# -----
# Filter for the highest version of Setup.exe
$setupLinks = $exeLinks | Where-Object { $_ -match "^https://www\.stationmaster\.com/Download/Setup\d+\.exe$" }

$highestVersion = 0
$downloadLink = $null

foreach ($link in $setupLinks) {
    $version = [regex]::Match($link, "Setup(\d+)\.exe").Groups[1].Value -as [int]
    if ($version -gt $highestVersion) {
        $highestVersion = $version
        $downloadLink = $link
    }
}

# Part 4 - Check size and hash to prevent redundant downloads with User-Agent
# PartVersion 1.02
# -----
if ($downloadLink) {
    # Create an HTTP request with User-Agent header to avoid 403 error
    $request = [System.Net.HttpWebRequest]::Create($downloadLink)
    $request.Method = "HEAD"
    $request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
    
    try {
        $response = $request.GetResponse()
        $contentLength = $response.ContentLength
        $response.Close()
    } catch {
        Write-Host "Error fetching response from server: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    # Extract the original filename from the download link
    $originalFilename = $downloadLink.Split('/')[-1]

    # Create a directory if it doesn't exist
    $downloadDirectory = "C:\winsm\SmartOffice_Installer"
    if (-not (Test-Path $downloadDirectory)) {
        New-Item -ItemType Directory -Path $downloadDirectory
    }

    # Determine the destination path
    $destinationPath = Join-Path -Path $downloadDirectory -ChildPath $originalFilename

    # Check if a file of the same size exists in the destination folder
    $existingFiles = Get-ChildItem -Path $downloadDirectory -Filter "*.exe"
    $fileExists = $existingFiles | Where-Object { $_.Length -eq $contentLength }

    # Download the file if no matching size file is found
    if (-not $fileExists) {
        Write-Host "Downloading new version of the Setup.exe..." -ForegroundColor Green
        Invoke-WebRequest -Uri $downloadLink -OutFile $destinationPath
    } else {
        Write-Host "File already exists and is up to date." -ForegroundColor Yellow
    }

    # Add timestamp to the downloaded file
    $timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
    $originalFilenameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($originalFilename)
    $extension = [System.IO.Path]::GetExtension($originalFilename)
    $newFileName = "${originalFilenameWithoutExtension}_${timestamp}${extension}"
    $newFilePath = Join-Path -Path $downloadDirectory -ChildPath $newFileName
    Rename-Item -Path $destinationPath -NewName $newFileName
}

# Part 5 - Delete older downloads
# PartVersion 1.00
# -----
$downloadedFiles = Get-ChildItem -Path $downloadDirectory -Filter "*.exe" | Sort-Object LastWriteTime -Descending
if ($downloadedFiles.Count -gt 1) {
    $filesToDelete = $downloadedFiles | Select-Object -Skip 1
    foreach ($file in $filesToDelete) {
        Remove-Item -Path $file.FullName -Force
    }
}
