# Invoke-WinGetOperation.ps1
A PowerShell wrapper for WinGet, supporting install, uninstall and list operations.

This script wraps around the winget.exe binary to provide enhanced usability for install, uninstall and list (detection) operations via Intune or SCCM. Additionally, this script logs all transactions to disk, colour-codes errors and extracts install failure exit codes and provides it to the caller, plus more.

For Version 2.0, we'll (probably) swap to WinGet's PowerShell module when its out of alpha/beta and if they do end up supporting PowerShell/WMF 5.1.
