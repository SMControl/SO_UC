# SO Upgrade Assistant
## Overview
- Automates the PRE and POST install tasks for SO Setup
- Works with both Current and Test versions
- Non-Destructive - Always safe to try assuming a backup has been performed
- Assistant Only - Does not interfere with the actual Setup, just automates the before and after

![Select Version](https://github.com/SMControl/SO_UC/blob/main/test/SOUA_SelectVersion.png)

## Task Performed
- Uses `SO_UC.exe` to download latest Current & Testing Setup versions (See more below)
- Installs Firebird fully with required settings and firebird.conf automatically if needed
- Handles Starting & Stopping LiveSales Service before and after
- Handles Starting & Stopping PDTWiFi & PDTWiFi64 before and after
- Sets folder permissions
## SO_UC.exe
- Schedules a daily task to run itself between 01:00 and 05:00
- Checks website for newer versions of Setup files
- Will only download files if a newer version is found; so usually does nothing.
## SmartOffice_Upgrade_Assistant.exe
- Copy of this program stored in `C:\winsm` for future ease of use
## Caveats
- Does not currently support Setup requiring reboots. (Either quit Assistant program and continue OR reboot PC after completion)
- Should not be used on non-standard installation paths. Anywhere where the install path is NOT `C:\Program Files (x86)`
## Requirements
- Windows 7 SP1 and newer
- Internet access for downloading latest Setup files
- Both  `SO_UC.exe` & `SmartOffice_Upgrade_Assistant.exe` must be allowed through the Firewall to function
## Usage
After a backup has been performed; run the following command in an Admin PowerShell:
```
irm https://raw.githubusercontent.com/SMControl/SO_UC/main/soua.ps1 | iex
```
Future usage can be done the same way or by running `C:\winsm\SmartOffice_Upgrade_Assistant.exe`
