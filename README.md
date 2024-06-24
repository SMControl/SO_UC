# SO Upgrade Assistant

## Overview

- Script automates the PRE and POST upgrade tasks for SO Setup.
- It does not install SO Automatically, it just automates some tasks.
- Always do a backup first.
- During script SO_UC.exe will require access through the firewall to retreive the latest SO Installer exe.
- Does not currently support SO Installer needing a reboot.

![](https://github.com/SMControl/SO_UC/blob/main/2024-06-24_21-29.png)


## Task Performed
1. Downloads latest Installer from SM and sets a task to check for new versions daily.
3. Checks Firebird; Installs if missing.
4. Deals with Firebird using programs and services PRE and POST Upgrade.
5. Waits for single instance of Firebird before launching the Setup.
6. Sets folder permissions

## Requirements

- Windows 7 SP1 or Windows Server 2008 R2 SP1
- Internet access for downloading latest Installer

## Usage

Run the following command in an Admin PowerShell to execute the script:
```
irm https://raw.githubusercontent.com/SMControl/SO_UC/main/soua.ps1 | iex
```

