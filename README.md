# SO Upgrade Assistant

## Overview

- Script automates the PRE and POST upgrade tasks for SO Setup.
- It does *NOT* install SO Automatically.
- Always do a backup first before running this script or manualy Installing/Upgrading.
- During script SO_UC.exe will require access through the firewall to retreive the latest SO Installer exe.
- Does not currently support Smart Office Installer needing a reboot.

![](https://github.com/SMControl/SO_UC/blob/main/2024-06-24_1101_1.png)


## Task Performed
1. Downloads latest Installer from SM
2. Sets a scheduled task to check nightly, between 00:00 and 05:59, for newer setup versions.
3. Checks Firebird; Installs with SM settings if missing.
4. Pre and Post Upgrade deals with Firebird using programs and services.
5. Waits for single instance of Firebird before launching the Setup.
6. Sets folder permissions for SM & Firebird 

## Requirements

- Windows 7 SP1 or Windows Server 2008 R2 SP1
- Internet access for downloading components

## Usage

Run the following command in an Admin PowerShell to execute the script:
```
irm https://raw.githubusercontent.com/SMControl/SO_UC/main/soua.ps1 | iex
```

