Write-Host "SO_UC_TEST.ps1 - Version 1.05"
# -----
# - Download progress, speed and eta

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

# Part 4 - Enhanced Download with Integer Percentages and Detailed ETA
# PartVersion 1.05
# -----
if ($downloadLink) {
    # Get the size of the file at the download link
    $request = [System.Net.HttpWebRequest]::Create($downloadLink)
    $request.Method = "HEAD"
    $response = $request.GetResponse()
    $contentLength = $response.ContentLength
    $response.Close()

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

    if (-not $fileExists) {
        # Initialize HttpClient and HttpRequestMessage for downloading
        $httpClient = [System.Net.Http.HttpClient]::new()
        $httpRequest = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $downloadLink)
        $httpRequest.Headers.Add("Accept-Encoding", "gzip, deflate")

        # Define variables for progress tracking
        $startTime = Get-Date
        $response = $httpClient.SendAsync($httpRequest, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).Result
        $totalBytes = $response.Content.Headers.ContentLength
        $stream = $response.Content.ReadAsStreamAsync().Result
        $fileStream = [System.IO.File]::Create($destinationPath)

        # Buffer for reading data
        $buffer = New-Object byte[] 8192
        $bytesRead = 0
        $downloadedBytes = 0

        # Define table header
        $header = "Downloading File: $originalFilename"
        Write-Host "`n$header"
        Write-Host "------------------------------------------"

        while (($bytesRead = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $fileStream.Write($buffer, 0, $bytesRead)
            $downloadedBytes += $bytesRead
            $elapsedTime = (Get-Date) - $startTime
            $speed = $downloadedBytes / $elapsedTime.TotalSeconds
            $remainingBytes = $totalBytes - $downloadedBytes
            $remainingTime = $remainingBytes / $speed

            # Update progress table
            $percentage = [math]::Round(($downloadedBytes / $totalBytes) * 100)
            $speedMB = [math]::Round($speed / 1MB, 2)
            $etaMinutes = [math]::Floor($remainingTime / 60)
            $etaSeconds = [math]::Round($remainingTime % 60)

            $progressLine = "{0,-35} {1,3}% {2,10} MB/s {3,2}m {4,2}s" -f "Progress:", $percentage, $speedMB, $etaMinutes, $etaSeconds
            Write-Host "`r$progressLine" -NoNewline -ForegroundColor Green
        }

        # Clean up
        $fileStream.Close()
        $stream.Close()
        $httpClient.Dispose()

        # Add timestamp to the downloaded file
        $timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
        $originalFilenameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($originalFilename)
        $extension = [System.IO.Path]::GetExtension($originalFilename)
        $newFileName = "${originalFilenameWithoutExtension}_${timestamp}${extension}"
        $newFilePath = Join-Path -Path $downloadDirectory -ChildPath $newFileName
        Rename-Item -Path $destinationPath -NewName $newFileName

        Write-Host "`rDownload completed and file renamed to $newFileName" -ForegroundColor Green
    } else {
        Write-Host "File with the same size already exists, no download needed." -ForegroundColor Yellow
    }
}

# Part 5 - delete older downloads
# PartVersion 1.00
# -----
$downloadedFiles = Get-ChildItem -Path $downloadDirectory -Filter "*.exe" | Sort-Object LastWriteTime -Descending
if ($downloadedFiles.Count -gt 1) {
    $filesToDelete = $downloadedFiles | Select-Object -Skip 1
    foreach ($file in $filesToDelete) {
        Remove-Item -Path $file.FullName -Force
    }
}
