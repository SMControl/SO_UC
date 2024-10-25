# SO Upgrade Assistant

## Overview

- Script automates the PRE and POST upgrade tasks for SO Setup.
- During script SO_UC.exe will require access through the firewall to retreive the latest SO Installer exe.
- Does not currently support SO Installer needing a reboot. (In this case an upgrade would be finished manually)

## Task Performed
1. Downloads latest Stable and Development Installers from SM and sets a task to check for new versions daily using **SO_UC.exe** (See below)
3. Checks Firebird; Installs with SM requirements if missing.
4. Deals with Firebird using programs and services PRE and POST Upgrade.
5. Waits for single instance of Firebird before launching the Setup.
6. Will ask which version of Setup to use, so suitable for installing both Stable & Development builds.
7. Sets folder permissions

## SO_UC.exe

- Sets a scheduled task to run itself at a random time between 01:00 and 05:00.
- Checks SM for newer version of Stable and Development Setup files and downloads them only if different than what it has.
- Stores setup files in "C:\winsm\SmartOffice_Installer"

## Requirements

- Windows 7 SP1 and newer
- Internet access for downloading latest Installer

## Usage

Run the following command in an Admin PowerShell to execute the script:
```
irm https://raw.githubusercontent.com/SMControl/SO_UC/main/soua.ps1 | iex
```

