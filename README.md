# SO Upgrade Assistant

## Overview

- Script automates the PRE and POST install tasks for SO Setup.
- Works with Current or Test versions and will obtain them.
- Non-Destructive - Always safe to try assuming a backup has been performed.
- Assistant Only - Does not interfere with the actual Setup, just automates the before and after.

![Select Version](https://github.com/SMControl/SO_UC/blob/main/test/SOUA_SelectVersion.png)
  
## Task Performed

- Uses `SO_UC.exe` to download latest Current & Testing Setup versions (See more below).
- Installs Firebird as required if needed.
- Before and After takes care of LivesSales Service and the two PDTWiFi's.
- Sets folder permissions

![OK to Finish](https://github.com/SMControl/SO_UC/blob/main/test/SOUA_OktoFinish.png)

![OK to Start](https://github.com/SMControl/SO_UC/blob/main/test/SOUA_OktoStart.png)

## SO_UC.exe

- Schedules a daily task to run itself between 01:00 and 05:00.
- Checks website for newer versions of Setup files.
- Will only download files if a newer version is found.

## SmartOffice_Upgrade_Assistant.exe
- Executable of this script stored in `C:\winsm` for future ease of use.

## Caveats

- Does not currently support Setup requiring reboots. (Either quit Assistant script and continue or reboot after complete.)
- Should not be used on non-standard installation paths. (Where install path is not `C:\Program Files (x86)`

## Requirements

- Windows 7 SP1 and newer
- Internet access for downloading latest Installer
- Both  `SO_UC.exe` & `SmartOffice_Upgrade_Assistant.exe` must be allowed through the Firewall to function.

## Usage

After a backup has been performed; run the following command in an Admin PowerShell:
```
irm https://raw.githubusercontent.com/SMControl/SO_UC/main/soua.ps1 | iex
```
Future runs can be done the same way or by running `C:\winsm\SmartOffice_Upgrade_Assistant.exe`
