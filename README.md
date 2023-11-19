# Get-PIDataArchiveUpgradeInfo
Monitor the health of your PI Data Archive with PowerShell

## Description 
This script generates CSV and TXT files with the following data:
- Archives List
- Critical Messages on Log
- Error Messages on Log
- Network Manager Statistics
- Tuning Parameters
- Windows System Info
- PI Services (with status and dependencies)
- OSIsoft Programs Installed on the Server
- Stale PI Points
- Bad PI Points

The script also generates an HTML file containing:
- AFLink Health Status
- Collective Members (if it is a collective)
- License Info
- Archive Info (including Archive Gaps list)
- Backup Info

## Prerequisites
- PowerShell Tools for the PI System, installed with the PI System Management Tools 2015 (3.5.1.7) or later

## Notes
This script is considerably faster if run directly on the PI Data Archive node. You must run this script as a local administrator on the PI Data Archive. Powershell must be run as Administrator as well. Your PowerShell Script Execution Policy must allow for the execution of this script.