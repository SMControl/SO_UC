# SO Upgrade Assistant

## Overview

- Script automates the PRE and POST upgrade/install tasks for SO Setup.
- During script SO_UC.exe will require access through the firewall to retreive the latest SO Installer exe. (One-time firewall pop-up)
- Does not currently support Setup requiring reboots. (Manually finish Upgrade or possibly reboot after)

## Task Performed

1. Downloads latest Stable & Development Setup executables from SM and sets a task to check for new versions daily using **SO_UC.exe** (See below)
2. Checks if Firebird is installed and does so with SM criteria if required.
3. Deals with Firebird using programs and services PRE and POST Upgrade (LiveSales, PDTWiFi etc).
5. Will ask which version of Setup to use, so suitable for installing/upgrading both Stable & Development builds.
6. Sets correct folder permissions.

## SO_UC.exe

- SO Update Checker
- Schedules a daily task to run itself at a random time between 01:00 and 05:00.
- Checks SM for newer version of Stable and Development Setup files.
- Will only download files if a newer version is found.
- Stores setup files in "winsm\SmartOffice_Installer"

## Requirements

- Windows 7 SP1 and newer
- Internet access for downloading latest Installer
- SO_UC.exe & SmartOffice_Upgrade_Assistant.exe must be allowed through the Firewall

## Usage

After a backup has been performed; run the following command in an Admin PowerShell:
```
irm https://raw.githubusercontent.com/SMControl/SO_UC/main/soua.ps1 | iex
```
