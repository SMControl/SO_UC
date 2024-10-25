# SO Upgrade Assistant

## Overview

- Script automates the PRE and POST upgrade/install tasks for SO Setup.
- During script SO_UC.exe will require access through the firewall to retreive the latest SO Installer exe. (Only really a one time pop-up issue with Norton etc)
- Does not currently support Setup requiring reboots.

## Task Performed

1. Downloads latest Stable and Development Installers from SM and sets a task to check for new versions daily using **SO_UC.exe** (See below)
2. Checks if Firebird is installed and does so with SM requirements if needed.
3. Deals with Firebird using programs and services PRE and POST Upgrade (LiveSales, PDTWiFi etc).
4. Waits for single instance of Firebird before launching the Setup.
6. Will ask which version of Setup to use, so suitable for installing/upgrading both Stable & Development builds.
7. Sets correct folder permissions and cleans up.

## SO_UC.exe

- SO Update Checker
- Sets a scheduled task to run itself at a random time between 01:00 and 05:00.
- Checks SM for newer version of Stable and Development Setup files and downloads them only if they are new.
- Stores setup files in "winsm\SmartOffice_Installer"

## Requirements

- Windows 7 SP1 and newer
- Internet access for downloading latest Installer

## Usage

Run the following command in an Admin PowerShell to execute the script:
```
irm https://raw.githubusercontent.com/SMControl/SO_UC/main/soua.ps1 | iex
```

