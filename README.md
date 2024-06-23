# SO Upgrade Assistant

## Overview

- Script automates the PRE and POST upgrade tasks for SO Setup.
- It does *NOT* install SO Automatically.
- Always do a backup first before running this script or manualy Installing/Upgrading.
- During script SO_UC.exe will require access through the firewall to retreive the latest SO Installer exe.

![Screenshot of Smart Office Upgrade Client](https://github.com/SMControl/SO_UC/raw/main/2024-06-23_1230.png)

## Requirements

- Windows 7 SP1 or Windows Server 2008 R2 SP1
- Internet access for downloading components

## Usage

Run the following command in an Admin PowerShell to execute the script:
```
irm https://raw.githubusercontent.com/SMControl/SO_UC/main/soua.ps1 | iex
```

