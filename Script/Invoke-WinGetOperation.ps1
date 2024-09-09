
<#PSScriptInfo

.VERSION 1.27

.GUID 74f34e90-5e24-4f61-b52e-facac43afaa9

.AUTHOR Mitch Richters

.COPYRIGHT Copyright © 2024 Mitch Richters. All rights reserved.

.TAGS winget

.LICENSEURI https://opensource.org/license/bsd-3-clause

.PROJECTURI https://github.com/mjr4077au/Invoke-WinGetOperation.ps1

.RELEASENOTES
- Initial public release for MSFT user group presentation.

#>

<#

.SYNOPSIS
A PowerShell wrapper for WinGet, supporting install, uninstall and list operations.

.DESCRIPTION
This script wraps around the winget.exe binary to provide enhanced usability for install, uninstall and list (detection) operations via Intune or SCCM.

Additionally, this script logs all transactions to disk, colour-codes errors and extracts install failure exit codes and provides it to the caller, plus more.

For Version 2.0, we'll (probably) swap to WinGet's PowerShell module when its out of alpha/beta and if they do end up supporting PowerShell/WMF 5.1.

.PARAMETER Install
Installs a WinGet package.

.PARAMETER Uninstall
Uninstalls a WinGet Package.

.PARAMETER Id
The ID of the package for the requested operation.

.PARAMETER Version
The Version of the package for the requested operation (if null, the latest will be installed).

.PARAMETER Scope
Whether the package is machine or user-scoped (if null, the packaged will be scoped to the machine).

.PARAMETER Source
Provide the pre-configured WinGet source where the package should come from.

.PARAMETER Installer-Type
Provide the type of installer that WinGet should use. Some manifests give both MSI and non-MSI installer types.

.PARAMETER Architecture
The archtecture of the package for the requested operation (if null, WinGet's default behaviour will be used).

.PARAMETER Custom
Install/uninstall arguments to append to default arguments during the requested operation.

.PARAMETER Override
Install/uninstall arguments to override default arguments during the requested operation.

.PARAMETER Force
Forces an install/uninstall operation to be attempted, even if the application is deemed installed/uninstalled already.

.PARAMETER IgnoreHashFailure
Allows overriding a failed installer hash comparison in an administrative context. Useful for apps like Google Chrome where the URLs are not versioned.

.PARAMETER DebugHashFailure
Forces the $IgnoreHashFailure pathway for debugging purposes without having to edit the script to force it.

.EXAMPLE
powershell.exe -ExecutionPolicy Bypass -NoProfile -NonInteractive -File Invoke-WinGetOperation.ps1 -NoLogo -Id Google.Chrome

[2023-07-11T09:36:11] Invoke-WinGetOperation.ps1 1.11
[2023-07-11T09:36:11] Transcript started, output file is C:\Users\mrichters\AppData\Local\ScriptLogs\Invoke-WinGetOperation.ps1\Invoke-WinGetOperation.ps1_list_Google.Chrome_20230711-093609.log
[2023-07-11T09:36:12] Using winget path: C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_1.20.1572.0_x64__8wekyb3d8bbwe\winget.exe
[2023-07-11T09:36:12] Executing: winget.exe list --exact --accept-source-agreements --Id Google.Chrome
[2023-07-11T09:36:14] Successfully detected Google Chrome [Google.Chrome] 114.0.5735.199.
[2023-07-11T09:36:14] Script finished with exit code: 0

.EXAMPLE
powershell.exe -ExecutionPolicy Bypass -NoProfile -NonInteractive -File Invoke-WinGetOperation.ps1 -NoLogo -Install -Id Microsoft.VSTOR

[2023-07-07T14:51:04] Invoke-WinGetOperation.ps1 1.11
[2023-07-07T14:51:04] Using winget path: C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_1.20.1572.0_x64__8wekyb3d8bbwe\winget.exe
[2023-07-07T14:51:04] Transcript started, output file is C:\Users\mrichters\AppData\Local\ScriptLogs\Invoke-WinGetOperation.ps1\Invoke-WinGetOperation.ps1_uninstall_Microsoft.VSTOR_20230707-145104.log
[2023-07-07T14:51:04] Executing: winget.exe install --exact --accept-source-agreements --accept-package-agreements --silent --id Microsoft.VSTOR --scope Machine
[2023-07-07T14:51:06] No applicable installer found, attempting to execute again without '--scope' argument.
[2023-07-07T14:51:06] Executing: winget.exe install --exact --accept-source-agreements --accept-package-agreements --silent --id Microsoft.VSTOR
[2023-07-07T14:51:07] Found Microsoft Visual Studio Tools for Office Runtime 2010 Redistributable [Microsoft.VSTOR] Version 10.0.60828.00.
[2023-07-07T14:51:07] Downloading https://download.microsoft.com/download/C/A/8/CA86DFA0-81F3-4568-875A-7E7A598D4C1C/vstor_redist.exe.
[2023-07-07T14:51:17] Successfully verified installer hash.
[2023-07-07T14:51:17] Starting package install...
[2023-07-07T14:51:36] Successfully installed.
[2023-07-07T14:51:36] Script finished with exit code: 0

.EXAMPLE
powershell.exe -ExecutionPolicy Bypass -NoProfile -NonInteractive -File Invoke-WinGetOperation.ps1 -NoLogo -Uninstall -Id Microsoft.VSTOR

[2023-07-07T14:51:04] Invoke-WinGetOperation.ps1 1.11
[2023-07-07T14:51:04] Using winget path: C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_1.20.1572.0_x64__8wekyb3d8bbwe\winget.exe
[2023-07-07T14:51:04] Transcript started, output file is C:\Users\mrichters\AppData\Local\ScriptLogs\Invoke-WinGetOperation.ps1\Invoke-WinGetOperation.ps1_uninstall_Microsoft.VSTOR_20230707-145104.log
[2023-07-07T14:51:04] Executing: winget.exe uninstall --exact --accept-source-agreements --accept-package-agreements --silent --id Microsoft.VSTOR --scope Machine
[2023-07-07T14:51:17] Starting package uninstall...
[2023-07-07T14:51:36] Successfully uninstalled.
[2023-07-07T14:51:36] Script finished with exit code: 0

.EXAMPLE
powershell.exe -ExecutionPolicy Bypass -NoProfile -NonInteractive -File Invoke-WinGetOperation.ps1 -NoLogo -Install -Id Microsoft.VSTOR -IgnoreHashFailure

[2023-07-11T20:31:37] Invoke-WinGetOperation.ps1 1.12
[2023-07-11T20:31:38] Using winget path: C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_1.20.1572.0_x64__8wekyb3d8bbwe\winget.exe
[2023-07-11T20:31:39] Executing: winget.exe install --exact --accept-source-agreements --accept-package-agreements --silent --Id Microsoft.VSTOR --Scope Machine
[2023-07-11T20:31:41] No applicable installer found, attempting to execute again without '--scope' argument.
[2023-07-11T20:31:41] Executing: winget.exe install --exact --accept-source-agreements --accept-package-agreements --silent --id Microsoft.VSTOR
[2023-07-11T20:31:43] Installation failed due to mismatched hash, attempting to override as '$IgnoreHashFailure' has been passed.
[2023-07-11T20:31:49] Installing/updating NuGet package provider, please wait...
[2023-07-11T20:31:52] Installing/updating psyml module, please wait...
[2023-07-11T20:32:05] Importing psyml module, please wait...
[2023-07-11T20:32:09] Obtained and processed package manifest from GitHub.
[2023-07-11T20:32:09] Downloading https://download.microsoft.com/download/C/A/8/CA86DFA0-81F3-4568-875A-7E7A598D4C1C/vstor_redist.exe.
[2023-07-11T20:33:10] Starting package install...
[2023-07-11T20:33:38] Successfully installed.
[2023-07-11T20:33:38] Script finished with exit code: 0

.INPUTS
None. You cannot pipe objects to Invoke-WinGetOperation.ps1.

.OUTPUTS
stdout stream. Invoke-WinGetOperation.ps1 returns a log string via Write-Host that can be piped.
stderr stream. Invoke-WinGetOperation.ps1 writes all error text to stderr for catching externally to PowerShell if required.

.NOTES
Tags: winget
Website: https://github.com/mjr4077au/
Copyright: (c) 2024 Mitchell Richters, BSD 3-Clause License
License: https://opensource.org/license/bsd-3-clause

**Changelog**

1.26.2
- Handle exit codes coming through from parsed manifests.
- Get rid of the delayed printing stuff, it doesn't work well with PSADT.

1.26.1
- Fix formatting issue with WinGet errors when there's no translation in the exit code lookup table.

1.26
- Clean up some parameter setups.
- Remove awkward try/catch setup in `Get-WinGetAppArguments`.
- Replace all `-join` operators with `[System.String]::Join()` static method.
- Revert `$wgAppVersion` changes from 84746649dcab0ad016c9b9ff1caf09483663c2c2.
- Remove some unnecessary version casts.
- Fix potential issue in `Get-WinGetAppInstaller` if specifying an installer type and the installer has no type specified.

1.25.2
- Remove some unnecessary argument passing.

1.25.1
- No need to statically provide the script's name to `Open-LogFile`.
- Add missing default switch-case to `Get-DefaultKnownSwitches`.

1.25
- Add missing #Requires.
- Remove `$action` global.
- Restore log output from `Initialize-NugetModule`.
- Remove remnants of `-ListAll` pathway.
- Add ability to debug $IgnoreHashFailire pathway.

1.24
- Remove `-ListAll` action.
- Remove `-ListExitCodes` action.
- Leverage repaired `Get-RedirectedUri` from backend to fix sideloading of hash-failed apps.

1.23
- Add support for `--installer-type`, handy for packages available as MSI and non-MSI.
- Add support for nested installers within `-IgnoreHashFailure` mode.
- Repair logic used when choosing an installer during `-IgnoreHashFailure` when there's multiple choices.
- Repair `Get-RedirectedUrl` to make it work properly, and make it PowerShell 7.x-compliant.
- Ensure `Write-LogEntry` also captures lines written to stderr to the transcript.
- Update `Get-ErrorRecord` to the latest iteration.
- Update `Write-StdErrMessage` to the latest iteration.

1.22
- Repaired call to `Get-WinGetAppManifest` where the parameter `-Version` was not renamed to `-AppVersion`.

1.21
- Add some level of detection for missing MSVC++ Runtime.
- Change `Get-WinGetAppManifest` `$Version` to `$AppVersion` to avoid collisions with the global namespace.
- Change script parameter `$Version` from System.String to System.Version so that leading/trailing zeros can be preserved.
- Change `Get-WinGetAppManifest` `$AppVersion` from System.String to System.Version so that leading/trailing zeros can be preserved.
- Change `Convert-WinGetListOutput`'s version parsing from System.String to System.Version so that leading/trailing zeros can be preserved.

1.20
- Re-saved script as UTF-8 but without byte-order mark (BOM).
- Updated `Initialize-NugetModule` to the latest iteration.
- Updated `Out-FriendlyErrorMessage` to the latest iteration.

1.19
- Removed pre-processing of error messages, it shouldn't be needed with `Out-FriendlyErrorMessage`.
- Old processing line: $_.Exception.Message.Split("`n")[0] -creplace '^(.+)(?<!\.)\.?At\s.+','$1.'
- Remove `Convert-WinGetShowOutput`. We really just need the version number and its proven unreliable.
- Do better testing for null string parameters that get passed through to `Convert-ParamsToArgArray`.

1.18
- Ensure `Get-WinGetAppInstaller` factors in the system locale for app with multiple languages.
- Rework `Get-WinGetAppInstaller` to handle situations where apps have multiple installer types.
- Fix double post++ on variable where a single ++pre will work.
- Allow `Write-LogEntry` to accept pipelined messages.

1.17
- Rework installer switch setup for `-IgnoreHashFailure` mode. The mode will now try six different ways to determine silent switches for installation.
- Add `Get-RedirectedUri` to parse URLs in the manifest for their absolute download location. so we can reliably extract the filename.
- Add bumpers before each function declaration to break the code up. The script is over 1000 lines now and needed it.
- Add `Out-FriendlyErrorMessage` to provide rich information for all errors we don't throw ourselves.
- Remove unused variable `$errMessage` from main catch block.
- Suppress `PSPossibleIncorrectUsageOfAssignmentOperator` warning from `Invoke-ScriptAnalyzer`.
- Do additional null check in `Convert-ParamsToArgArray`'s final `Where-Object` call.
- Optimise loop in `Out-WinGetExitCodeTable`.

1.16
- Repair issue where `$Source` argument wasn't permitted against `list` operations.
- Repair issue where trailing periods weren't trimmed in `Convert-WinGetListOutput`.

1.15
- Default $Source to 'winget' to speed up processing.

1.14
- Reworked `Initialize-NugetModule` to only test Nuget state when we absolutely have to.
- For `-IgnoreHashFailure`, properly factor in `-Custom` and/or `-Override` if it's passed.
- For `-IgnoreHashFailure`, test whether arguments is null before calling `Start-Process`.
- For `-IgnoreHashFailure`, repair issue where Start-Process wasn't using $PWD as working directory.
- For `-IgnoreHashFailure`, set shell's $ProgressPreference to Ignore to speed up downloads.

1.13
- Minor optimisation to `Convert-WinGetListOutput`.
- Update `Initialize-NugetModule` to newest version.
- Fixed case-insensitivty in catch block regex parsing.
- Fixed typo in WinGet version testing.
- Bumped minimum version from 1.4.10173 to 1.5.1572. Update your pre-reqs!
- Reversed change "Filter out some noisy WinGet logging".
- Reversed change "Add `--verbose` to winget calls if specified on the script".
- Enabled verbose WinGet logging by default.
- Added `-Force` argument to pass through to WinGet.
- Added `-Custom` argument to pass through to WinGet.
- Added `-Source` argument to pass through to WinGet.
- Added error handling for invalid sources.
- Rework log filename to "ScriptName_Action_Id_DateTime_{0}".
- Formatted script transaction log filename as "Script".
- Formatted winget action output log filename as "WinGet".
- Added `--log` to WinGet param array, with logging to aforementioned file.
- Renamed `-ForceOnHashFailure` to `-IgnoreHashFailure`.
- Added `-ListExitCodes` to print table of exit codes WinGet may exit with.
- Slight improvement to error reporting.

1.12
- Fixed warnings from PSScriptAnalyzer. Violations of `PSAvoidUsingWriteHost` are deliberate, `PSReviewUnusedParameter` warnings are a bug: https://github.com/PowerShell/PSScriptAnalyzer/issues/1472
- Added (experimental) workaround for apps failing signature checks due to use of rolling release URIs.
- Dynamically source all exit codes from WinGet's source code to improve error output.
- Update URL provided to user to refer to when WinGet operation fails.
- Advise of the exit code the script exits with in the log.

1.11
- Add proper conversion of 'winget.exe list' output.
- Fix bad regex that had non-escaped period.
- Add $action to generated log filename.
- Add proper catching of terminating errors.
- Add `--verbose` to winget calls if specified on the script.
- Add mode to use script as a method of getting full `winget.exe list` output as objects.
- Filter out some noisy WinGet logging.
- Throw if scope is user but we're running as system.
- Throw if winget.exe is not able to be found.
- Throw if winget.exe is less than minimum version.
- Throw if console is somehow not UTF-8.
- Added colouring to error lines.
- Added ability to colour lines for warnings.
- Added ability to colour lines for success.
- Added colons to the timestamp output for a certain someone 🙃.
- Added information about the script version to the log.
- Fixed minor issue affecting operation in PowerShell 7.x environments.
- Fixed log directory test to not only rely on scope, but also user's permissions.
- Improve log output by only writing the last line from winget.exe if its informational.
- Added super-awesome ANSI art title banner, just because.

1.10
- Changed PowerShell's output encoding to UTF-8 to match WinGet.
- Removed verbose logging and changed all to stdout.
- Added scope testing for apps where the manifest does not define a scope, therefore isn't found.
- Improved overall logging by not discarding full WinGet logging.
- Re-write exit code with app installer's exit code where possible.

1.9
- Changed log handling to test the scope prior to execution.

1.8
- Add verbose output if the caller requests it.

1.7
- Removed defaulting for `$Architecture` value so the architecture isn't forced by default.

1.6
- Overhaul internals of `Out-WinGetArgs` and how it creates an argument array from the script's parameters.

1.5
- Re-add `$Scope` argument to script. Some apps default to user context 🤦‍♂.

1.4
- Break out code into functions for easier readability.
- Add handling for version when installing/uninstalling an app.
- Add specialised handling for version when detecting an app.
- Remove `$Scope` argument as the script is intended for system usage only.

1.3
- Rework argument handling.

1.2
- Configure script to transcribe all output to log file.

1.1
- Rework handling of winget path detection for when running in SYSTEM context.

1.0
- Initial release!

#>

[CmdletBinding(DefaultParameterSetName = 'list')]
param
(
	[Parameter(Mandatory = $true, ParameterSetName = 'install')]
	[System.Management.Automation.SwitchParameter]$Install,

	[Parameter(Mandatory = $true, ParameterSetName = 'uninstall')]
	[System.Management.Automation.SwitchParameter]$Uninstall,

	[Parameter(Mandatory = $true, ParameterSetName = 'list', HelpMessage = 'WinGet Argument')]
	[Parameter(Mandatory = $true, ParameterSetName = 'install', HelpMessage = 'WinGet Argument')]
	[Parameter(Mandatory = $true, ParameterSetName = 'uninstall', HelpMessage = 'WinGet Argument')]
	[ValidateNotNullOrEmpty()]
	[System.String]$Id,

	[Parameter(Mandatory = $false, ParameterSetName = 'list', HelpMessage = 'WinGet Argument')]
	[Parameter(Mandatory = $false, ParameterSetName = 'install', HelpMessage = 'WinGet Argument')]
	[Parameter(Mandatory = $false, ParameterSetName = 'uninstall', HelpMessage = 'WinGet Argument')]
	[ValidatePattern('^\d+(?=\.)[\d.]+$')]
	[System.String]$Version,  # Must be a string to preserve leading/trailing zeros!

	[Parameter(Mandatory = $false, ParameterSetName = 'install', HelpMessage = 'WinGet Argument')]
	[Parameter(Mandatory = $false, ParameterSetName = 'uninstall', HelpMessage = 'WinGet Argument')]
	[ValidateSet('User','Machine')]
	[System.String]$Scope = 'Machine',

	[Parameter(Mandatory = $false, ParameterSetName = 'list', HelpMessage = 'WinGet Argument')]
	[Parameter(Mandatory = $false, ParameterSetName = 'install', HelpMessage = 'WinGet Argument')]
	[Parameter(Mandatory = $false, ParameterSetName = 'uninstall', HelpMessage = 'WinGet Argument')]
	[ValidateNotNullOrEmpty()]
	[System.String]$Source = 'winget',

	[Parameter(Mandatory = $false, ParameterSetName = 'install', HelpMessage = 'WinGet Argument')]
	[ValidateSet('Burn','Wix','Msi','Nullsoft','Inno')]
	[Alias('InstallerType')]
	[System.String]${Installer-Type},

	[Parameter(Mandatory = $false, ParameterSetName = 'install', HelpMessage = 'WinGet Argument')]
	[ValidateSet('x86','x64')]
	[System.String]$Architecture,

	[Parameter(Mandatory = $false, ParameterSetName = 'install', HelpMessage = 'WinGet Argument')]
	[ValidateNotNullOrEmpty()]
	[System.String]$Custom,

	[Parameter(Mandatory = $false, ParameterSetName = 'install', HelpMessage = 'WinGet Argument')]
	[ValidateNotNullOrEmpty()]
	[System.String]$Override,

	[Parameter(Mandatory = $false, ParameterSetName = 'install', HelpMessage = 'WinGet Argument')]
	[Parameter(Mandatory = $false, ParameterSetName = 'uninstall', HelpMessage = 'WinGet Argument')]
	[System.Management.Automation.SwitchParameter]$Force,

	[Parameter(Mandatory = $false, ParameterSetName = 'install')]
	[System.Management.Automation.SwitchParameter]$IgnoreHashFailure,

	[Parameter(Mandatory = $false, ParameterSetName = 'install')]
	[System.Management.Automation.SwitchParameter]$DebugHashFailure
)


#---------------------------------------------------------------------------
#
# Critical global variables required for script functionality.
#
#---------------------------------------------------------------------------

# Set required variables to ensure script functionality.
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
$ProgressPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
Set-PSDebug -Strict
Set-StrictMode -Version Latest

# Set console encoding as required for winget.exe output.
$lastConsoleOutputEncoding = [System.Console]::OutputEncoding
[System.Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Define global script properties.
$script = @{
	Command = $MyInvocation.MyCommand
	Action = $PSCmdlet.ParameterSetName
	NonInteractive = ![System.Environment]::UserInteractive -or [System.Environment]::GetCommandLineArgs().Contains('-NonInteractive')
	User = [System.Security.Principal.WindowsIdentity]::GetCurrent()
	MinVersion = [System.Version]::new(1, 7, 10582)
	ExitCode = 0
}


#---------------------------------------------------------------------------
#
# Test whether user has administrative rights, or their session is UAC-elevated.
#
#---------------------------------------------------------------------------

function Test-ElevatedSession
{
	param
	(
		[ValidateNotNullOrEmpty()]
		[System.Security.Principal.WindowsIdentity]$Identity = $script.User
	)

	return [System.Security.Principal.WindowsPrincipal]::new($Identity).IsInRole([System.Security.Principal.WindowsBuiltinRole]::Administrator)
}


#---------------------------------------------------------------------------
#
# Writes a string to stderr. By default, PowerShell writes everything to stdout, even errors.
#
#---------------------------------------------------------------------------

function Write-StdErrMessage
{
	begin
	{
		# If we're in a ConsoleHost, colour the output like Write-Error does.
		if ($consoleHost = $Host.Name.Equals('ConsoleHost'))
		{
			[System.Console]::BackgroundColor = [System.ConsoleColor]::Black
			[System.Console]::ForegroundColor = [System.ConsoleColor]::Red
		}
	}

	process
	{
		# We can only truly write to stderr in a ConsoleHost.
		if ($consoleHost)
		{
			[System.Console]::Error.WriteLine($_)
		}
		else
		{
			$Host.UI.WriteErrorLine($_)
		}
	}

	end
	{
		# Reset the colour back to defaults.
		if ($consoleHost)
		{
			[System.Console]::ResetColor()
		}
	}
}


#---------------------------------------------------------------------------
#
# Converts an array of objects into a single bulleted string with line feeds.
#
#---------------------------------------------------------------------------

function ConvertTo-BulletedList
{
	return "`n$([System.String]::Join("`n", ($input.Where({![System.String]::IsNullOrWhiteSpace($_)}) -replace "^","> ")))"
}


#---------------------------------------------------------------------------
#
# Returns an InnerException's ErrorRecord if the piped ErrorRecord has one.
#
#---------------------------------------------------------------------------

filter Get-ErrorRecord
{
	# Return piped object if we don't have an InnerException to process.
	if (!($InnerException = $_.Exception.InnerException))
	{
		return $_
	}

	# Return piped object if we don't have an InnerException ErrorRecord to process.
	if (!$InnerException.PSObject.Properties.Name.Contains('ErrorRecord'))
	{
		return $_
	}

	# Return piped object if the InnerException ErrorRecord is null.
	if (!($InnerError = $InnerException.ErrorRecord))
	{
		return $_
	}

	# Return piped object if the InnerException ErrorRecord's stack trace is null.
	if ([System.String]::IsNullOrWhiteSpace($InnerError.ScriptStackTrace))
	{
		return $_
	}

	# If we've made it this far, the InnerException's ErrorRecord is the favourable ErrorRecord.
	return $InnerError
}


#---------------------------------------------------------------------------
#
# Converts an ErrorRecord's ScriptStackTrace into an array of objects.
#
#---------------------------------------------------------------------------

filter Convert-StackTraceToObjectArray
{
	# Split the stack trace on each new line, then process each line into an object.
	$arr = $_.Split("`n").Trim().ForEach({
		$arr = [System.Text.RegularExpressions.Regex]::Match($_, '^at\s(.+),\s(.+):\sline\s(\d+)$').Groups.Value[1..3]
		return [pscustomobject][ordered]@{Function = $arr[0]; Path = $arr[1]; Line = [System.UInt32]$arr[2]}
	})

	# Add a ToString() method prior to returning. This gives the caller the original line, effectively.
	$method = {"at $($this.Function), $($this.Path): line $($this.Line)"}
	return $arr | Microsoft.PowerShell.Utility\Add-Member -MemberType ScriptMethod -Name ToString -Value $method -PassThru -Force
}


#---------------------------------------------------------------------------
#
# Converts an ErrorRecord into an object with data for use with `Out-FriendlyErrorMessage`.
#
#---------------------------------------------------------------------------

filter Out-FriendlyErrorData
{
	# Process the stacktrace into an object array and set up the index. We use a combination of lines to return the best data.
	$stackData = $_.ScriptStackTrace | Convert-StackTraceToObjectArray; $i = $stackData[0].ToString() -match '<.+, <'
	$errCmdlet = $_.InvocationInfo.MyCommand | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty Name
	$ipoCmdlet = ($stackData[$i+1].Function -match '^Invoke-PipelineOperation') -and !$stackData[$i].Path.Equals($Script:MyInvocation.MyCommand.ScriptBlock.Module.Path)

	# Process the error message. Some strings like ones from a ApplicationFailedException/NativeCommandFailed include a partial stack trace.
	if (($errMsgStr = $_.Exception.Message) -match "At (line:|[A-Z]:\\)")
	{
		$errMsgStr = [System.String]::Join($null, $errMsgStr.Split("`n")) -replace '^(.+)At (line:|[A-Z]:\\).+$','$1'
	}

	# Build out our return object, performing some tests along the way.
	return [pscustomobject][ordered]@{
		Function = ($func = $stackData[$(if ($stackData[$i].Function.StartsWith('<')) {$i+1+$ipoCmdlet} else {$i})].Function)
		Command = $(if ($ipoCmdlet -and ($errCmdlet -eq 'ForEach-Object')) {$stackData[$i+1].Function} else {$errCmdlet}) | Microsoft.PowerShell.Core\Where-Object {!$_.StartsWith($func)}
		Message = "$($errMsgStr.Split("`n").Trim().TrimEnd('.').Where({![System.String]::IsNullOrWhiteSpace($_)}) -replace '([^,;:])$','$1.')"
		Path = $(if (!$stackData[$i].Path.StartsWith('<')) {[System.IO.Path]::GetDirectoryName($stackData[$i].Path)})
		File = $(if (!$stackData[$i].Path.StartsWith('<')) {[System.IO.Path]::GetFileName($stackData[$i].Path)})
		Line = $stackData[$i].Line
	}
}


#---------------------------------------------------------------------------
#
# Returns a pretty multi-line error message derived from the provided ErrorRecord.
#
#---------------------------------------------------------------------------

filter Out-FriendlyErrorMessage
{
	param
	(
		[ValidateNotNullOrEmpty()]
		[System.String]$ErrorPrefix
	)

	# If there's no custom error prefix, add a generic one.
	if ([System.String]::IsNullOrWhiteSpace($ErrorPrefix)) {$ErrorPrefix = 'ERROR:'}

	# Get our error data in preparation for formatting.
	$errData = $_ | Out-FriendlyErrorData
	$errCmdl = if ($errData.Command) {"$($errData.Command): "}
	$stckArr = $_.ScriptStackTrace.Split("`n").Trim().Where({$_ -notmatch '^at <.+>, <.+>: line 1$'})
	$stckMsg = if ($stckArr.Count -gt 1) {" Full stack trace:$($stckArr -replace '^at ' | ConvertTo-BulletedList)"}

	# Return our formatted string, factoring in some null checks also.
	return [System.String]::Join("`n", "$ErrorPrefix $($errData.File): Line #$($errData.Line): $($errData.Function): $errCmdl$($errData.Message)$stckMsg".Split("`n").Trim())
}


#---------------------------------------------------------------------------
#
# Displays our banner and copyright, then commences transcribing to disk.
#
#---------------------------------------------------------------------------

function Open-LogFile
{
	# Only draw the ANSI art for interactive sessions.
	if (!$script.NonInteractive)
	{
		# Set up strings for printing. They're encoded to ensure the spacing and line breaks are correct.
		$titleArt = [System.Text.Encoding]::GetEncoding(437).GetString([System.Convert]::FromBase64String('CiAgICAg29u7ICAgINvbu9vbu9vb27sgICDb27sg29vb29vbuyDb29vb29vbu9vb29vb29vbuyAgCiAgICAg29u6ICAgINvbutvbutvb29u7ICDb27rb28nNzc3NvCDb28nNzc3NvMjNzdvbyc3NvCAgCiAgICAg29u6INu7INvbutvbutvbydvbuyDb27rb27ogINvb27vb29vb27sgICAgINvbuiAgICAgCiAgICAg29u629vbu9vbutvbutvbusjb27vb27rb27ogICDb27rb28nNzbwgICAgINvbuiAgICAgCiAgICAgyNvb28nb29vJvNvbutvbuiDI29vb27rI29vb29vbybzb29vb29vbuyAgINvbuiAgICAgCiAgICAgIMjNzbzIzc28IMjNvMjNvCAgyM3NzbwgyM3Nzc3NvCDIzc3Nzc3NvCAgIMjNvCAgICAgCgog29vb29vb29u7INvb29vb27sgINvb29vb27sg29u7ICAgICDb27sgINvbu9vbu9vb29vb29vbuwogyM3N29vJzc2829vJzc3N29u729vJzc3N29u729u6ICAgICDb27og29vJvNvbusjNzdvbyc3NvAogICAg29u6ICAg29u6ICAg29u629u6ICAg29u629u6ICAgICDb29vb28m8INvbuiAgINvbuiAgIAogICAg29u6ICAg29u6ICAg29u629u6ICAg29u629u6ICAgICDb28nN29u7INvbuiAgINvbuiAgIAogICAg29u6ICAgyNvb29vb28m8yNvb29vb28m829vb29vb27vb27ogINvbu9vbuiAgINvbuiAgIAogICAgyM28ICAgIMjNzc3NzbwgIMjNzc3NzbwgyM3Nzc3NzbzIzbwgIMjNvMjNvCAgIMjNvCAgIAo=')) -split "`n"
		$subtitle = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('IOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlAogV2luR2V0IHdyYXBwZXIgZm9yIEludHVuZSBpbnN0YWxsL3JlbW92ZSBhbmQgZGV0ZWN0aW9ucwogQ29weXJpZ2h0IMKpIDIwMjQgTWl0Y2hlbGwgUmljaHRlcnMsIEFsbCByaWdodHMgcmVzZXJ2ZWQKIOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlOKAlAo='))

		# Back up the console cursor state and disable it.
		$cursorState = [System.Console]::CursorVisible
		[System.Console]::CursorVisible = $false

		# Output ANSI art to console and restore console. If condition is true, do a raster scan, otherwise just vertically line-by-line.
		# Ensure this always uses the .ForEach() method to avoid introduced pipeline latency.
		if ($true)
		{
			# Get each line.
			$titleArt.ForEach({
				# Print each line's character, then place a line feed before continuing.
				$_.GetEnumerator().ForEach({
					# Log the start time, write the character and spin until enough ticks have elapsed. 10000 ticks is 1 millisecond.
					# Note: Halting the thread via Start-Sleep or [System.Threading.Thread]::Sleep() is not precise enough here.
					$start = [System.DateTime]::Now.Ticks
					[System.Console]::Write($_)
					while (([System.DateTime]::Now.Ticks - $start) -lt 20000) {}
				})
				[System.Console]::WriteLine()
			})
		}
		else
		{
			# Get each line and draw one-by-one.
			$titleArt.ForEach({
				[System.Console]::WriteLine($_)
				[System.Threading.Thread]::Sleep(125)
			})
		}
		[System.Console]::CursorVisible = $cursorState

		# Draw subtitle.
		[System.Console]::WriteLine($subtitle)
		[System.Threading.Thread]::Sleep(2000)
	}

	# Set up all log vars and commence transcription of all output to disk.
	$logBase = @("$Env:LOCALAPPDATA\ScriptLogs", "$Env:SYSTEMROOT\Logs")[(Test-ElevatedSession) -and ($Scope -eq 'Machine')]
	$logPath = "$([System.IO.Directory]::CreateDirectory($logBase).FullName)\$($script.Command.Name)"
	$logFile = "$($script.Command.Name)_$($script.Action)_$($Id)_$([System.DateTime]::Now.ToString('yyyyMMdd-hhmmss'))_{0}.log"
	$logStart = (Start-Transcript -Path (($script.LogFile = "$logPath\$logFile") -f 'Script')) -replace '(^Transcript started, output file is )(.+)$', '$1[$2].'
	Write-LogEntry -Message "$($script.Command.Name) $((Test-ScriptFileInfo -LiteralPath $script.Command.Source).Version)"
	Write-LogEntry -Message $logStart -Type StdOut
	return $true
}


#---------------------------------------------------------------------------
#
# Advise on the success of our operation and close out log file.
#
#---------------------------------------------------------------------------

function Close-LogFile
{
	# Advise of the exit code used before closing out the log.
	Write-LogEntry -Message "Script finished with exit code [$($script.ExitCode)]."
	$null = Stop-Transcript
}


#---------------------------------------------------------------------------
#
# Writes a timestamped log entry to the console.
#
#---------------------------------------------------------------------------

filter Write-LogEntry
{
	param
	(
		[Parameter(Mandatory = $true, ValueFromPipeline = $true)]
		[ValidateNotNullOrEmpty()]
		[System.String]$Message,

		[Parameter(Mandatory = $false)]
		[ValidateSet('Error', 'Warning', 'Success', 'StdOut')]
		[System.String]$Type
	)

	# Generate output message. Using `[System.DateTime]::Now` static method is _much_ faster than `Get-Date`.
	$msg = "[$([System.DateTime]::Now.ToString('yyyy-MM-ddTHH:mm:ss'))] $Message"

	# Handle the message based on the type.
	switch ($Type)
	{
		'Error'
		{
			# We need to use Write-Host here to capture it in a transcript.
			$msg | Write-StdErrMessage
			$msg | Write-Host 6>$null
			break
		}
		'Warning'
		{
			# Colour the output just like Write-Warning does, then reset.
			$msg | Write-Host -ForegroundColor Yellow -BackgroundColor Black
			break
		}
		'Success'
		{
			$msg | Write-Host -ForegroundColor Green -BackgroundColor Black
			break
		}
		'StdOut'
		{
			[System.Console]::WriteLine($msg)
		}
		default
		{
			$msg | Write-Host
			break
		}
	}
}


#---------------------------------------------------------------------------
#
# For the given ParameterMetadata, convert it into a native executable's argument list.
#
#---------------------------------------------------------------------------

function Convert-ParamsToArgArray
{
	param
	(
		[Parameter(Mandatory = $true, ParameterSetName = 'Preset')]
		[Parameter(Mandatory = $true, ParameterSetName = 'Custom')]
		[ValidateNotNullOrEmpty()]
		[System.Management.Automation.ParameterMetadata[]]$Parameters,

		[Parameter(Mandatory = $true, ParameterSetName = 'Preset')]
		[Parameter(Mandatory = $true, ParameterSetName = 'Custom')]
		[ValidateNotNullOrEmpty()]
		[System.String]$ParameterSetName,

		[Parameter(Mandatory = $true, ParameterSetName = 'Preset')]
		[Parameter(Mandatory = $true, ParameterSetName = 'Custom')]
		[ValidateNotNullOrEmpty()]
		[System.String]$HelpMsgTag,

		[Parameter(Mandatory = $true, ParameterSetName = 'Preset')]
		[ValidateSet('MSI', 'Winget', 'DellCommandUpdate')]
		[System.String]$Preset,

		[Parameter(Mandatory = $true, ParameterSetName = 'Custom')]
		[ValidateSet(' ', '=', "`n")]
		[System.String]$ArgValSeparator,

		[Parameter(Mandatory = $false, ParameterSetName = 'Custom')]
		[ValidateSet('-', '--', '/')]
		[System.String]$ArgPrefix,

		[Parameter(Mandatory = $false, ParameterSetName = 'Custom')]
		[ValidateSet("'", '"')]
		[System.String]$ValueWrapper,

		[Parameter(Mandatory = $false, ParameterSetName = 'Preset')]
		[Parameter(Mandatory = $false, ParameterSetName = 'Custom')]
		[ValidateSet(',','|')]
		[System.String]$MultiValDelimiter = ','
	)

	# Set up the string for formatting.
	$string = switch ($Preset) {
		'MSI'               {"{0}=`"{1}`""}
		'Winget'            {"--{0}`n{1}"}
		'DellCommandUpdate' {"-{0}={1}"}
		default             {"$($ArgPrefix){0}$($ArgValSeparator)$($ValueWrapper){1}$($ValueWrapper)"}
	}

	# Set up the param filter, it's quite long.
	$filter = {
		($_.ParameterSets.ContainsKey($ParameterSetName)) -and
		($ParamAttribute = $_.Attributes | Where-Object {$_ -is [System.Management.Automation.ParameterAttribute]}) -and
		($ParamAttribute.HelpMessage -eq $HelpMsgTag)
	}

	# Set up process script.
	$process = {
		if (($Value = Get-Variable -Name $_.Name -ValueOnly) -and (($Value -isnot [System.String]) -or ![System.String]::IsNullOrWhiteSpace($Value))) {
			$string -f $_.Name,$(if ($Value -isnot [System.Management.Automation.SwitchParameter]) {$Value -join $MultiValDelimiter}) -split "`n"
		}
	}

	# Return an array of args based on our filter and process, filtering out any nulls.
	return $Parameters | Where-Object $filter | ForEach-Object $process | Where-Object {![System.String]::IsNullOrWhiteSpace($_)}
}


#---------------------------------------------------------------------------
#
# Parses a URI for its actual download location.
#
#---------------------------------------------------------------------------

function Get-RedirectedUri
{
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[System.Uri]$Uri
	)

	# Create web request.
	$webReq = [System.Net.WebRequest]::Create($Uri)
	$webReq.AllowAutoRedirect = $false
	$webReq.Accept = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7'

	# Get a response and close it out.
	$reqRes = $webReq.GetResponse()
	$resLoc = $reqRes.GetResponseHeader('Location')
	$reqRes.Close()

	# If $resLoc is empty, return the provided URI so something is returned to the caller.
	if ([System.String]::IsNullOrWhiteSpace($resLoc))
	{
		return $Uri
	}
	return [System.Uri]$resLoc
}


#---------------------------------------------------------------------------
#
# Parses a URI for its file name.
#
#---------------------------------------------------------------------------

function Get-UriFileName
{
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[System.Uri]$Uri
	)

	# Re-write the URI to factor in any redirections.
	$Uri = Get-RedirectedUri -Uri $Uri
	$badChars = "($([System.String]::Join('|', [System.IO.Path]::GetInvalidFileNameChars().ForEach({[System.Text.RegularExpressions.Regex]::Escape($_)}))))"

	# Create web request.
	$webReq = [System.Net.WebRequest]::Create($Uri)
	$webReq.AllowAutoRedirect = $false
	$webReq.Accept = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7'

	# Get a response and close it out.
	$reqRes = $webReq.GetResponse()
	$resCnt = $reqRes.GetResponseHeader('Content-Disposition')
	$reqRes.Close()

	# If $resCnt is empty, the provided URI likely has the filename in it.
	if (!$resCnt.Contains('filename'))
	{
		return $Uri.ToString().Split('/')[-1] -replace $badChars
	}
	return $resCnt.Split(';').Trim().Where({$_.StartsWith('filename=')}).Split('=').Trim()[-1] -replace $badChars
}


#---------------------------------------------------------------------------
#
# Imports a Nuget module, optionally downloading it and any package management dependencies.
#
#---------------------------------------------------------------------------

function Initialize-NuGetModule
{
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[System.String[]]$Name,

		[Parameter(Mandatory = $false, HelpMessage = 'PackageManagement parameter')]
		[ValidateNotNullOrEmpty()]
		[System.Version]$MinimumVersion,

		[Parameter(Mandatory = $false, HelpMessage = 'PackageManagement parameter')]
		[ValidateNotNullOrEmpty()]
		[System.Version]$MaximumVersion,

		[Parameter(Mandatory = $false, HelpMessage = 'PackageManagement parameter')]
		[ValidateNotNullOrEmpty()]
		[System.Version]$RequiredVersion,

		[Parameter(Mandatory = $false, HelpMessage = 'PackageManagement parameter')]
		[ValidateSet('AllUsers', 'CurrentUser')]
		[System.String]$Scope
	)

	# Define internal functions for repeated usage.
	function Install-NuGetModule
	{
		param
		(
			[Parameter(Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[System.String]$ModuleName
		)

		PowerShellGet\Install-Module @imParams -Name $ModuleName -Force -Confirm:$false -Verbose:$false
	}
	function Uninstall-NuGetModule
	{
		param
		(
			[Parameter(Mandatory = $true)]
			[ValidateNotNullOrEmpty()]
			[System.Object[]]$Modules
		)

		$Modules | Microsoft.PowerShell.Core\Remove-Module -Force -Confirm:$false -Verbose:$false
		$Modules | PowerShellGet\Uninstall-Module -Force -Confirm:$false -Verbose:$false
	}

	# Amend the scope if it's not been provided.
	if (!$Scope)
	{
		$PSBoundParameters.Add('Scope', ($Scope = if (Test-ElevatedSession) {'AllUsers'} else {'CurrentUser'}))
	}

	# Perform initial function setup.
	Write-LogEntry -Message "Importing PackageManagement module, please wait..."
	Microsoft.PowerShell.Core\Import-Module -Name PackageManagement -Verbose:$false

	# Ensure the NuGet package provider is installed on the system.
	Write-LogEntry -Message "Confirming NuGet package provider state, please wait..."
	if (!($provider = PackageManagement\Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction Ignore -Verbose:$false) -or ($provider.Version -lt [System.Version]::new(2, 8, 5, 201)))
	{
		Write-LogEntry -Message "Installing/updating NuGet package provider, please wait..."
		[System.Void](PackageManagement\Install-PackageProvider -Name NuGet -Scope $Scope -Force -Verbose:$false)
	}

	# Cache module install state. Calling PowerShellGet\Get-InstalledModule is slow so let's only do it once.
	Write-LogEntry -Message "Gathering installed module information, please wait..."
	$modules = PowerShellGet\Get-InstalledModule -Name $Name -AllVersions -ErrorAction Ignore -Verbose:$false

	# Set up variables needed within loop.
	$null = $PSBoundParameters.Remove('Name')
	$imParams = $PSBoundParameters
	$docspath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::MyDocuments)
	$elevated = $Scope.Equals('AllUsers')

	# Loop through supplied names to determine whether they're installed or not.
	foreach ($item in $Name)
	{
		# Test whether module is installed at all, followed by any specific version requirements.
		if (!($module = $modules | Microsoft.PowerShell.Core\Where-Object {$_.Name.Equals($item)}))
		{
			Write-LogEntry -Message "Installing $item module, please wait..."
			Install-NuGetModule -ModuleName $item
		}
		else
		{
			# Handle versions that are too old/new.
			if ($MaximumVersion -and ($newer = $module | Microsoft.PowerShell.Core\Where-Object {($_.Version -gt $MaximumVersion) -and ($elevated -or $_.InstalledLocation.StartsWith($docspath))}))
			{
				Write-LogEntry -Message "Uninstalling $item versions newer than $MaximumVersion, please wait..."
				Uninstall-NuGetModule -Module $newer
				$module = $module | Microsoft.PowerShell.Core\Where-Object {$newer -notcontains $_}
			}
			if ($MinimumVersion -and ($older = $module | Microsoft.PowerShell.Core\Where-Object {($_.Version -lt $MinimumVersion) -and ($elevated -or $_.InstalledLocation.StartsWith($docspath))}))
			{
				Write-LogEntry -Message "Uninstalling $item versions older than $MinimumVersion, please wait..."
				Uninstall-NuGetModule -Module $older
				$module = $module | Microsoft.PowerShell.Core\Where-Object {$older -notcontains $_}
			}

			# Handle required vs. out of date versions.
			if (!$module -or ($RequiredVersion -and ($module.Version -ne $RequiredVersion)))
			{
				Write-LogEntry -Message "Installing $item module, please wait..."
				Install-NuGetModule -ModuleName $item
			}
			elseif ((!$RequiredVersion -or ($module.Version -ne $RequiredVersion)) -and (!$MaximumVersion -or ($module.Version -lt $MaximumVersion)) -and (($latest = PowerShellGet\Find-Module -Name $item -Verbose:$false).Version -gt $module.Version))
			{
				# Unfortunately, Update-Module _does not_ update versions, it merely installs newer versions side-by-side...
				# Before attempting removal, confirm we're able to do so. A user can't uninstall a system-installed module.
				if ($elevated -or $module.InstalledLocation.StartsWith($docspath))
				{
					Write-LogEntry -Message "Updating $item module to $($latest.Version), please wait..."
					Uninstall-NuGetModule -Module $module
				}
				else
				{
					Write-LogEntry -Message "Installing $item module, please wait..."
				}
				Install-NuGetModule -ModuleName $item
			}
		}

		# Import module if it's not currently loaded.
		if (!(Microsoft.PowerShell.Core\Get-Module).Name.Contains($item))
		{
			Write-LogEntry -Message "Importing $item module."
			Microsoft.PowerShell.Core\Import-Module -Name $item -Verbose:$false
		}
	}
}


#---------------------------------------------------------------------------
#
# Ensures our console has the correct encoding for winget.exe.
#
#---------------------------------------------------------------------------

function Test-ConsoleEncoding
{
	# Throw if the console isn't UTF-8. I tried setting and resetting the encoding in the functions
	# where this actually matters, but it did not work in PS7+ so a global setting is necessary...
	if (![System.Console]::OutputEncoding.Equals([System.Text.Encoding]::UTF8))
	{
		throw "This script requires a UTF-8 console for successful operation."
	}
}


#---------------------------------------------------------------------------
#
# Ensures we're not trying to install a user-scoped item into NT AUTHORITY\SYSTEM.
#
#---------------------------------------------------------------------------

function Test-WinGetScope
{
	# We need to throw if such a mistake in setup occurs as it leads to a wild goose chase when diagnosing what's gone wrong.
	if ($systemUser -and $Scope.Equals('User'))
	{
		throw "Installing user-scoped applications as the system user is not supported."
	}
}


#---------------------------------------------------------------------------
#
# Installs the latest Visual Studio Runtime dependency.
#
#---------------------------------------------------------------------------

function Install-VisualStudioRuntimeDependency
{
	# We can only reliably update if we're elevated.
	if (!(Test-ElevatedSession))
	{
		throw "The installed version of WinGet was unable to run. Please ensure the latest Visual Studio 2015-2022 Runtime is installed and try again."
	}

	# Set required variables for install operation.
	$pkgArch = @('x86','x64')[[System.Environment]::Is64BitProcess]
	$pkgName = "Microsoft Visual C++ 2015-2022 Redistributable ($pkgArch)"
	$uriPath = "https://aka.ms/vs/17/release/vc_redist.$pkgArch.exe"
	Write-LogEntry -Message "Preparing $pkgName dependency, please wait..."

	# Define arguments for installation.
	$spParams = @{
		FilePath = "$([System.IO.Path]::GetTempPath())$(Get-Random).exe"
		ArgumentList = "/install", "/quiet", "/norestart", "/log $($script.LogFile -f 'MSVCRT')"
	}

	# Download and extract installer.
	Write-LogEntry -Message "Downloading [$pkgName], please wait..."
	Microsoft.PowerShell.Utility\Invoke-WebRequest -UseBasicParsing -Uri $uriPath -OutFile $spParams.FilePath

	# Invoke installer.
	Write-LogEntry -Message "Installing [$pkgName], please wait..."
	if ($script.ExitCode = (Microsoft.PowerShell.Management\Start-Process @spParams -Wait -PassThru).ExitCode)
	{
		throw "The installation of [$pkgName] failed with exit code [$($script.ExitCode)]."
	}
}


#---------------------------------------------------------------------------
#
# Pre-provisions the latest WinGet binaries.
#
#---------------------------------------------------------------------------

function Install-DesktopAppInstallerDependency
{
	# We can only reliably update if we're elevated.
	if (!(Test-ElevatedSession))
	{
		throw "WinGet is not installed, or the installed version is less than $($script.MinVersion). Please update Microsoft.DesktopAppInstaller and try again."
	}

	# Update WinGet to the latest version. Don't rely in 3rd party store API services for this.
	# https://learn.microsoft.com/en-us/windows/package-manager/winget/#install-winget-on-windows-sandbox
	Write-LogEntry -Message "Updating $(($pkgName = "Microsoft.DesktopAppInstaller")) dependency, please wait..."

	# Define installation file info.
	$packages = @(
		@{
			Name = 'C++ Desktop Bridge Runtime dependency'
			Uri = ($uri = [System.Uri]'https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx')
			FilePath = "$($env:TEMP)\$($uri.Segments[-1])"
		}
		@{
			Name = 'Windows UI Library dependency'
			Uri = ($uri = [System.Uri]'https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.x64.appx')
			FilePath = "$($env:TEMP)\$($uri.Segments[-1])"
		}
		@{
			Name = 'latest WinGet msixbundle'
			Uri = ($uri = Get-RedirectedUri -Uri 'https://aka.ms/getwinget')
			FilePath = "$($env:TEMP)\$($uri.Segments[-1])"
		}
	)

	# Download all packages.
	foreach ($package in $packages)
	{
		Write-LogEntry -Message "Downloading [$($package.Name)], please wait..."
		Microsoft.PowerShell.Utility\Invoke-WebRequest -UseBasicParsing -Uri $package.Uri -OutFile $package.FilePath
	}

	# Pre-provision package in the system.
	$aappParams = @{
		Online = $true
		SkipLicense = $true
		PackagePath = $packages[(-1)].FilePath
		DependencyPackagePath = $packages[(0)..($packages.Count-2)].FilePath
		LogPath = $script.LogFile -f 'Dism'
	}
	Write-LogEntry -Message "Pre-provisioning [$pkgName] $($packages[-1].Uri.Segments[-2].Trim('/')), please wait..."
	$null = Dism\Add-AppxProvisionedPackage @aappParams
}


#---------------------------------------------------------------------------
#
# Gets the correct, fully-qualified path to winget.exe.
#
#---------------------------------------------------------------------------

function Out-WinGetPath
{
	# Internal function to get the WinGet Path. We can't rely on the output of Get-AppxPackage for some systems as it'll update, but Get-AppxPackage won't reflect the new path fast enough.
	function Get-WinGetPath
	{
		# For the system user, get the path from Program Files directly.
		if ($systemUser)
		{
			return (Get-ChildItem -Path "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller*\winget.exe" | Sort-Object -Descending | Select-Object -First 1)
		}
		elseif ([System.IO.File]::Exists(($wingetPath = "$(Appx\Get-AppxPackage -Name Microsoft.DesktopAppInstaller -AllUsers:$systemUser | Microsoft.PowerShell.Utility\Sort-Object -Property Version -Descending | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty InstallLocation -First 1)\winget.exe")))
		{
			return $wingetPath
		}
	}

	# Test whether WinGet is installed and available at all.
	if (!($wingetPath = Get-WinGetPath) -or ![System.IO.File]::Exists($wingetPath))
	{
		Install-DesktopAppInstallerDependency
		if (!($wingetPath = Get-WinGetPath) -or ![System.IO.File]::Exists($wingetPath))
		{
			throw "Failed to get a valid WinGet path after successfully pre-provisioning the app. Please report this issue for further analysis."
		}
	}

	# Test whether we have any output from winget.exe. If this is null, it typically means the appropriate MSVC++ runtime is not installed.
	if (!($wingetOutput = & $wingetPath))
	{
		Install-VisualStudioRuntimeDependency
		if (!($wingetOutput = & $wingetPath))
		{
			throw "The installed version of WinGet was unable to run. This is possibly related to the Visual Studio 2015-2022 Runtime."
		}
	}

	# Ensure winget.exe is above the minimum version.
	if ([System.Version](($wingetOutput | Microsoft.PowerShell.Utility\Select-Object -First 1) -replace '^.+\sv') -lt $script.MinVersion)
	{
		# Install the missing dependency and reset variables.
		Install-DesktopAppInstallerDependency
		$wingetPath = Get-WinGetPath

		# Ensure winget.exe is above the minimum version.
		if ([System.Version]($wingetVer = (($wingetOutput = & $wingetPath) | Microsoft.PowerShell.Utility\Select-Object -First 1) -replace '^.+\sv') -lt $script.MinVersion)
		{
			throw "The installed WinGet version of $wingetVer is less than $($script.MinVersion). Please check the DISM pre-provisioning logs and try again."
		}

		# Reset WinGet sources after updating. Helps with a corner-case issue discovered.
		Write-LogEntry -Message "Resetting all WinGet sources following update, please wait..."
		if (!($wgSrcRes = & $wingetPath source reset --force 2>&1).Equals('Resetting all sources...Done'))
		{
			Write-LogEntry -Message "An issue occurred while resetting WinGet sources [$($wgSrcRes.TrimEnd('.'))]. Continuing with operation." -Warning
		}
	}

	# Return tested path to the caller.
	Write-LogEntry -Message "Using WinGet path [$wingetPath]."
	return $wingetPath
}


#---------------------------------------------------------------------------
#
# Returns an argument array for splatting onto winget.exe.
#
#---------------------------------------------------------------------------

function Out-WinGetArgArray
{
	param
	(
		[ValidateNotNullOrEmpty()]
		[System.String[]]$Exclude
	)

	# Before doing anything, test we're not trying to pass custom arguments and override at the same time.
	if ($Custom -and $Override)
	{
		throw "The usage of '-Custom' and '-Override' at the same time is not supported."
	}

	# Define exclusions.
	$exclusions = $(if ($Exclude) {$Exclude}; if ($script.Action.Equals('list')) {'Version'})

	# Standard args.
	$script.Action
	'--exact'
	'--verbose-logs'
	'--accept-source-agreements'

	# Calculated args from function's parameter block.
	$cpaParams = @{
		Parameters = $Script:MyInvocation.MyCommand.Parameters.Values.Where({$exclusions -notcontains $_.Name})
		ParameterSetName = $script.Action
		HelpMsgTag = 'WinGet Argument'
		Preset = 'WinGet'
	}
	Convert-ParamsToArgArray @cpaParams

	# Calculated args based on the function's action.
	if ($script.Action.Equals('install'))
	{
		'--accept-package-agreements'
	}
	if (!$script.Action.Equals('list'))
	{
		'--silent'
		'--log'; $script.LogFile -f 'WinGet'
	}
}


#---------------------------------------------------------------------------
#
# Converts the results of `winget.exe list` into PowerShell PSCustomObjects.
#
#---------------------------------------------------------------------------

function Convert-WinGetListOutput
{
	<#

	.NOTES
	This function expects the console to be UTF-8 using `[System.Console]::OutputEncoding = [System.Text.Encoding]::UTF8`.

	The encoding is globally set in this function, but exporting this function and using it without setting the console first will have unexpected results.

	Attempts were made to set this and reset it in in the begin{} and end{} blocks respectively, but this only worked with PowerShell 5.1 and not 7.x.

	#>

	begin
	{
		# Define variables for heading data that'll be the first line via the pipe.
		$listHeading = $headIndices = $null
	}

	process
	{
		# Filter out nonsense lines from the pipe.
		if (($line = $_.Trim().TrimEnd('.')) -notmatch '^\w+')
		{
			return
		}

		# Use our first valid line to set up the keys for each property.
		if (!$listHeading)
		{
			# Get all headings and the indices from the output.
			$listHeading = $line -split '\s+'
			$headIndices = $listHeading.ForEach({$line.IndexOf($_)}) + 10000
			return
		}

		# Establish hashtable to hold contents we're converting.
		$obj = [ordered]@{}

		# Begin conversion and return object to the pipeline.
		for ($i = 0; $i -lt $listHeading.Length; $i++)
		{
			$thisi = [System.Math]::Min($headIndices[$i], $line.Length)
			$nexti = [System.Math]::Min($headIndices[$i+1], $line.Length)
			$value = $line.Substring($thisi, $nexti-$thisi).Trim()
			$obj.Add($listHeading[$i], $(if (![System.String]::IsNullOrWhiteSpace($value)) {$value}))
		}
		return [pscustomobject]$obj
	}
}


#---------------------------------------------------------------------------
#
# Invokes `winget.exe` with the specified arguments, processing the plain text output throug the pipeline.
#
#---------------------------------------------------------------------------

function Invoke-WinGetExecutable
{
	<#

	.NOTES
	This function expects the console to be UTF-8 using `[System.Console]::OutputEncoding = [System.Text.Encoding]::UTF8`.

	The encoding is globally set in this function, but exporting this function and using it without setting the console first will have unexpected results.

	Attempts were made to set this and reset it in in the begin{} and end{} blocks respectively, but this only worked with PowerShell 5.1 and not 7.x.

	#>

	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[System.String]$LiteralPath,

		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[System.String[]]$Arguments,

		[Parameter(Mandatory = $false)]
		[System.Management.Automation.SwitchParameter]$Silent
	)

	# Test whether the specified source is valid before continuing.
	if ($Source -and !($sources = & $LiteralPath source list | Convert-WinGetListOutput).Name.Contains($Source))
	{
		throw "The specified source '$Source' is not valid. Currently configured WinGet sources are [$([System.String]::Join(', ', $sources.Name -replace '^|$','"'))]."
	}

	# Invoke winget and print each non-null line that doesn't contain a hex code.
	Write-LogEntry -Message "Executing [winget.exe $Arguments]."
	$wgOutput = & $LiteralPath @Arguments | Microsoft.PowerShell.Core\Where-Object {$_ -match '^\w+'} | Microsoft.PowerShell.Core\ForEach-Object {
		($line = $_.Trim() -replace '((?<![.:])|:)$','.')
		if (!$Silent -and ($line -notmatch '^0x\w{8}')) {
			Write-LogEntry -Message $line
		}
	}

	# Return accumulated output to the caller. This must be a string array!
	return ,[System.String[]]$wgOutput
}


#---------------------------------------------------------------------------
#
# Munges all `winget.exe` exit codes from the application's source code into a PowerShell PSCustomObject.
#
#---------------------------------------------------------------------------

function Out-WinGetExitCodeTable
{
	param
	(
		[System.Management.Automation.SwitchParameter]$Dynamic
	)

	# If dynamic, get all WinGet exit codes from the source code and munge it into a PSCustomObject for lookups.
	if ($Dynamic)
	{
		$sourceUri = 'https://raw.githubusercontent.com/microsoft/winget-cli/master/src/AppInstallerSharedLib/Public/AppInstallerErrors.h'
		return (Microsoft.PowerShell.Utility\Invoke-RestMethod -UseBasicParsing -Uri $sourceUri).Split("`n").Where({$_ -match '^.+_ERROR_.+0x\w{8}'}).ForEach({
			begin
			{
				$obj = [ordered]@{}
			}

			process
			{
				$obj.Add(($_ -replace '^.+_ERROR_(\w+).+$','$1'), [System.Int32]($_ -replace '^.+(0x\w{8}).+$','$1'))
			}

			end
			{
				if ($obj.Count) {return [pscustomobject]$obj}
			}
		})
	}

	# If we're here, return a hard coded list in case generating the below lookup table fails on unstable internet connections.
	return [pscustomobject]@{
		INTERNAL_ERROR = -1978335231
		INVALID_CL_ARGUMENTS = -1978335230
		COMMAND_FAILED = -1978335229
		MANIFEST_FAILED = -1978335228
		CTRL_SIGNAL_RECEIVED = -1978335227
		SHELLEXEC_INSTALL_FAILED = -1978335226
		UNSUPPORTED_MANIFESTVERSION = -1978335225
		DOWNLOAD_FAILED = -1978335224
		CANNOT_WRITE_TO_UPLEVEL_INDEX = -1978335223
		INDEX_INTEGRITY_COMPROMISED = -1978335222
		SOURCES_INVALID = -1978335221
		SOURCE_NAME_ALREADY_EXISTS = -1978335220
		INVALID_SOURCE_TYPE = -1978335219
		PACKAGE_IS_BUNDLE = -1978335218
		SOURCE_DATA_MISSING = -1978335217
		NO_APPLICABLE_INSTALLER = -1978335216
		INSTALLER_HASH_MISMATCH = -1978335215
		SOURCE_NAME_DOES_NOT_EXIST = -1978335214
		SOURCE_ARG_ALREADY_EXISTS = -1978335213
		NO_APPLICATIONS_FOUND = -1978335212
		NO_SOURCES_DEFINED = -1978335211
		MULTIPLE_APPLICATIONS_FOUND = -1978335210
		NO_MANIFEST_FOUND = -1978335209
		EXTENSION_PUBLIC_FAILED = -1978335208
		COMMAND_REQUIRES_ADMIN = -1978335207
		SOURCE_NOT_SECURE = -1978335206
		MSSTORE_BLOCKED_BY_POLICY = -1978335205
		MSSTORE_APP_BLOCKED_BY_POLICY = -1978335204
		EXPERIMENTAL_FEATURE_DISABLED = -1978335203
		MSSTORE_INSTALL_FAILED = -1978335202
		COMPLETE_INPUT_BAD = -1978335201
		YAML_INIT_FAILED = -1978335200
		YAML_INVALID_MAPPING_KEY = -1978335199
		YAML_DUPLICATE_MAPPING_KEY = -1978335198
		YAML_INVALID_OPERATION = -1978335197
		YAML_DOC_BUILD_FAILED = -1978335196
		YAML_INVALID_EMITTER_STATE = -1978335195
		YAML_INVALID_DATA = -1978335194
		LIBYAML_ERROR = -1978335193
		MANIFEST_VALIDATION_WARNING = -1978335192
		MANIFEST_VALIDATION_FAILURE = -1978335191
		INVALID_MANIFEST = -1978335190
		UPDATE_NOT_APPLICABLE = -1978335189
		UPDATE_ALL_HAS_FAILURE = -1978335188
		INSTALLER_SECURITY_CHECK_FAILED = -1978335187
		DOWNLOAD_SIZE_MISMATCH = -1978335186
		NO_UNINSTALL_INFO_FOUND = -1978335185
		EXEC_UNINSTALL_COMMAND_FAILED = -1978335184
		ICU_BREAK_ITERATOR_ERROR = -1978335183
		ICU_CASEMAP_ERROR = -1978335182
		ICU_REGEX_ERROR = -1978335181
		IMPORT_INSTALL_FAILED = -1978335180
		NOT_ALL_PACKAGES_FOUND = -1978335179
		JSON_INVALID_FILE = -1978335178
		SOURCE_NOT_REMOTE = -1978335177
		UNSUPPORTED_RESTSOURCE = -1978335176
		RESTSOURCE_INVALID_DATA = -1978335175
		BLOCKED_BY_POLICY = -1978335174
		RESTAPI_INTERNAL_ERROR = -1978335173
		RESTSOURCE_INVALID_URL = -1978335172
		RESTAPI_UNSUPPORTED_MIME_TYPE = -1978335171
		RESTSOURCE_INVALID_VERSION = -1978335170
		SOURCE_DATA_INTEGRITY_FAILURE = -1978335169
		STREAM_READ_FAILURE = -1978335168
		PACKAGE_AGREEMENTS_NOT_ACCEPTED = -1978335167
		PROMPT_INPUT_ERROR = -1978335166
		UNSUPPORTED_SOURCE_REQUEST = -1978335165
		RESTAPI_ENDPOINT_NOT_FOUND = -1978335164
		SOURCE_OPEN_FAILED = -1978335163
		SOURCE_AGREEMENTS_NOT_ACCEPTED = -1978335162
		CUSTOMHEADER_EXCEEDS_MAXLENGTH = -1978335161
		MISSING_RESOURCE_FILE = -1978335160
		MSI_INSTALL_FAILED = -1978335159
		INVALID_MSIEXEC_ARGUMENT = -1978335158
		FAILED_TO_OPEN_ALL_SOURCES = -1978335157
		DEPENDENCIES_VALIDATION_FAILED = -1978335156
		MISSING_PACKAGE = -1978335155
		INVALID_TABLE_COLUMN = -1978335154
		UPGRADE_VERSION_NOT_NEWER = -1978335153
		UPGRADE_VERSION_UNKNOWN = -1978335152
		ICU_CONVERSION_ERROR = -1978335151
		PORTABLE_INSTALL_FAILED = -1978335150
		PORTABLE_REPARSE_POINT_NOT_SUPPORTED = -1978335149
		PORTABLE_PACKAGE_ALREADY_EXISTS = -1978335148
		PORTABLE_SYMLINK_PATH_IS_DIRECTORY = -1978335147
		INSTALLER_PROHIBITS_ELEVATION = -1978335146
		PORTABLE_UNINSTALL_FAILED = -1978335145
		ARP_VERSION_VALIDATION_FAILED = -1978335144
		UNSUPPORTED_ARGUMENT = -1978335143
		BIND_WITH_EMBEDDED_NULL = -1978335142
		NESTEDINSTALLER_NOT_FOUND = -1978335141
		EXTRACT_ARCHIVE_FAILED = -1978335140
		NESTEDINSTALLER_INVALID_PATH = -1978335139
		PINNED_CERTIFICATE_MISMATCH = -1978335138
		INSTALL_LOCATION_REQUIRED = -1978335137
		ARCHIVE_SCAN_FAILED = -1978335136
		PACKAGE_ALREADY_INSTALLED = -1978335135
		PIN_ALREADY_EXISTS = -1978335134
		PIN_DOES_NOT_EXIST = -1978335133
		CANNOT_OPEN_PINNING_INDEX = -1978335132
		MULTIPLE_INSTALL_FAILED = -1978335131
		MULTIPLE_UNINSTALL_FAILED = -1978335130
		NOT_ALL_QUERIES_FOUND_SINGLE = -1978335129
		PACKAGE_IS_PINNED = -1978335128
		PACKAGE_IS_STUB = -1978335127
		APPTERMINATION_RECEIVED = -1978335126
		DOWNLOAD_DEPENDENCIES = -1978335125
		DOWNLOAD_COMMAND_PROHIBITED = -1978335124
		SERVICE_UNAVAILABLE = -1978335123
		RESUME_ID_NOT_FOUND = -1978335122
		CLIENT_VERSION_MISMATCH = -1978335121
		INVALID_RESUME_STATE = -1978335120
		CANNOT_OPEN_CHECKPOINT_INDEX = -1978335119
		RESUME_LIMIT_EXCEEDED = -1978335118
		INVALID_AUTHENTICATION_INFO = -1978335117
		AUTHENTICATION_TYPE_NOT_SUPPORTED = -1978335116
		AUTHENTICATION_FAILED = -1978335115
		AUTHENTICATION_INTERACTIVE_REQUIRED = -1978335114
		AUTHENTICATION_CANCELLED_BY_USER = -1978335113
		AUTHENTICATION_INCORRECT_ACCOUNT = -1978335112
		NO_REPAIR_INFO_FOUND = -1978335111
		REPAIR_NOT_APPLICABLE = -1978335110
		EXEC_REPAIR_FAILED = -1978335109
		REPAIR_NOT_SUPPORTED = -1978335108
		ADMIN_CONTEXT_REPAIR_PROHIBITED = -1978335107
		SQLITE_CONNECTION_TERMINATED = -1978335106
		DISPLAYCATALOG_API_FAILED = -1978335105
		NO_APPLICABLE_DISPLAYCATALOG_PACKAGE = -1978335104
		SFSCLIENT_API_FAILED = -1978335103
		NO_APPLICABLE_SFSCLIENT_PACKAGE = -1978335102
		LICENSING_API_FAILED = -1978335101
		INSTALL_PACKAGE_IN_USE = -1978334975
		INSTALL_INSTALL_IN_PROGRESS = -1978334974
		INSTALL_FILE_IN_USE = -1978334973
		INSTALL_MISSING_DEPENDENCY = -1978334972
		INSTALL_DISK_FULL = -1978334971
		INSTALL_INSUFFICIENT_MEMORY = -1978334970
		INSTALL_NO_NETWORK = -1978334969
		INSTALL_CONTACT_SUPPORT = -1978334968
		INSTALL_REBOOT_REQUIRED_TO_FINISH = -1978334967
		INSTALL_REBOOT_REQUIRED_FOR_INSTALL = -1978334966
		INSTALL_REBOOT_INITIATED = -1978334965
		INSTALL_CANCELLED_BY_USER = -1978334964
		INSTALL_ALREADY_INSTALLED = -1978334963
		INSTALL_DOWNGRADE = -1978334962
		INSTALL_BLOCKED_BY_POLICY = -1978334961
		INSTALL_DEPENDENCIES = -1978334960
		INSTALL_PACKAGE_IN_USE_BY_APPLICATION = -1978334959
		INSTALL_INVALID_PARAMETER = -1978334958
		INSTALL_SYSTEM_NOT_SUPPORTED = -1978334957
		INSTALL_UPGRADE_NOT_SUPPORTED = -1978334956
		INVALID_CONFIGURATION_FILE = -1978286079
		INVALID_YAML = -1978286078
		INVALID_FIELD_TYPE = -1978286077
		UNKNOWN_CONFIGURATION_FILE_VERSION = -1978286076
		SET_APPLY_FAILED = -1978286075
		DUPLICATE_IDENTIFIER = -1978286074
		MISSING_DEPENDENCY = -1978286073
		DEPENDENCY_UNSATISFIED = -1978286072
		ASSERTION_FAILED = -1978286071
		MANUALLY_SKIPPED = -1978286070
		WARNING_NOT_ACCEPTED = -1978286069
		SET_DEPENDENCY_CYCLE = -1978286068
		INVALID_FIELD_VALUE = -1978286067
		MISSING_FIELD = -1978286066
		TEST_FAILED = -1978286065
		TEST_NOT_RUN = -1978286064
		GET_FAILED = -1978286063
		UNIT_NOT_INSTALLED = -1978285823
		UNIT_NOT_FOUND_REPOSITORY = -1978285822
		UNIT_MULTIPLE_MATCHES = -1978285821
		UNIT_INVOKE_GET = -1978285820
		UNIT_INVOKE_TEST = -1978285819
		UNIT_INVOKE_SET = -1978285818
		UNIT_MODULE_CONFLICT = -1978285817
		UNIT_IMPORT_MODULE = -1978285816
		UNIT_INVOKE_INVALID_RESULT = -1978285815
		UNIT_SETTING_CONFIG_ROOT = -1978285808
		UNIT_IMPORT_MODULE_ADMIN = -1978285807
		NOT_SUPPORTED_BY_PROCESSOR = -1978285806
	}
}


#---------------------------------------------------------------------------
#
# Little filter to get nested installer type if present.
#
#---------------------------------------------------------------------------

filter Get-WinGetInstallerType
{
	if ($_.PSObject.Properties.Name.Contains('NestedInstallerType'))
	{
		return $_.NestedInstallerType
	}
	elseif ($_.PSObject.Properties.Name.Contains('InstallerType'))
	{
		return $_.InstallerType
	}
}


#---------------------------------------------------------------------------
#
# For `-IgnoreHashFailure`, returns the manifest for the provided WinGet Id.
#
#---------------------------------------------------------------------------

function Get-WinGetAppManifest
{
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[System.String]$AppVersion
	)

	# Ensure we have psyml available and imported.
	Initialize-NuGetModule -Name psyml -MinimumVersion ([System.Version]::new(1, 0, 0))

	# Set up vars and get package manifest.
	$wgUriBase = "https://raw.githubusercontent.com/microsoft/winget-pkgs/master/manifests/{0}/{1}/{2}/{3}.installer.yaml"
	$wgPkgsUri = [System.String]::Format($wgUriBase, $Id.Substring(0,1).ToLower(), $Id.Replace('.','/'), $AppVersion, $Id)
	$wgManifest = Microsoft.PowerShell.Utility\Invoke-RestMethod -UseBasicParsing -Uri $wgPkgsUri | ConvertFrom-Yaml
	Write-LogEntry -Message "Downloaded and parsed package manifest from GitHub."
	return $wgManifest
}


#---------------------------------------------------------------------------
#
# For `-IgnoreHashFailure`, returns the correct installer from a WinGet manifest.
#
#---------------------------------------------------------------------------

function Get-WinGetAppInstaller
{
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[pscustomobject]$Manifest
	)

	# Get correct installation data from the manifest based on scope and system architecture.
	$systemArch = ('x86', 'x64')[[System.Environment]::Is64BitOperatingSystem]
	$nativeArch = $Manifest.Installers.Architecture -contains $systemArch
	$cultureName = [System.Globalization.CultureInfo]::CurrentUICulture.Name
	$wgInstaller = $Manifest.Installers | Microsoft.PowerShell.Core\Where-Object {
		(!$_.PSObject.Properties.Name.Contains('Scope') -or ($_.Scope -eq $Scope)) -and
		(!$_.PSObject.Properties.Name.Contains('InstallerLocale') -or ($_.InstallerLocale -eq $cultureName)) -and
		(!${Installer-Type} -or (($instType = $_ | Get-WinGetInstallerType) -and ($instType -eq ${Installer-Type}))) -and
		($_.Architecture.Equals($Architecture) -or ($haveArch = $_.Architecture -eq $systemArch) -or (!$haveArch -and !$nativeArch))
	}

	# Validate the output. The yoda notation is to keep PSScriptAnalyzer happy.
	if ($null -eq $wgInstaller)
	{
		# We found nothing and therefore can't continue.
		throw "Error occurred while processing installer metadata from the package's manifest."
	}
	elseif ($wgInstaller -is [System.Collections.IEnumerable])
	{
		# We got multiple values. Get all unique installer types from the metadata and check for uniqueness.
		if (!$wgInstaller.Count.Equals((($wgInstTypes = $wgInstaller | Get-WinGetInstallerType | Microsoft.PowerShell.Utility\Select-Object -Unique) | Microsoft.PowerShell.Utility\Measure-Object).Count))
		{
			# Something's gone wrong as we've got duplicate installer types.
			throw "Error determining correct installer metadata from the package's manifest."
		}

		# Installer types were unique, just return the first one and hope for the best.
		Write-LogEntry -Message "Found installer types ['$([System.String]::Join("', '", $wgInstTypes))']; using [$($wgInstTypes[0])] metadata."
		$wgInstaller = $wgInstaller | Microsoft.PowerShell.Core\Where-Object {($_ | Get-WinGetInstallerType).Equals($wgInstTypes[0])}
	}

	# Return installer metadata to the caller.
	Write-LogEntry -Message "Processed installer metadata from package manifest."
	return $wgInstaller
}


#---------------------------------------------------------------------------
#
# For `-IgnoreHashFailure`, downloads the installer and returns the file path to it.
#
#---------------------------------------------------------------------------

function Get-WinGetAppDownload
{
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[pscustomobject]$Installer
	)

	# Download WinGet app and store path to binary.
	$wgFilePath = "$([System.IO.Directory]::CreateDirectory("$([System.IO.Path]::GetTempPath())$(Get-Random)").FullName)\$(Get-UriFileName -Uri $Installer.InstallerUrl)"
	Write-LogEntry -Message "Downloading [$($Installer.InstallerUrl)], please wait..."
	Microsoft.PowerShell.Utility\Invoke-WebRequest -UseBasicParsing -Uri $Installer.InstallerUrl -OutFile $wgFilePath

	# If downloaded file is a zip, we need to expand it and modify our file path before returning.
	if ($wgFilePath -match 'zip$')
	{
		Write-LogEntry -Message "Downloaded installer is a zip file, expanding its contents."
		Microsoft.PowerShell.Archive\Expand-Archive -LiteralPath $wgFilePath -DestinationPath $Env:TEMP -Force
		$wgFilePath = "$($Env:TEMP)\$($Installer.NestedInstallerFiles.RelativeFilePath)"
	}
	return $wgFilePath
}


#---------------------------------------------------------------------------
#
# For `-IgnoreHashFailure`, returns all valid exit codes for the given installer.
#
#---------------------------------------------------------------------------

function Get-WinGetAppExitCodes
{
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[pscustomobject]$Manifest,

		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[pscustomobject]$Installer,

		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[System.String]$FilePath
	)

	# Try to get switches from the installer, then the manifest, then by whatever known defaults we have.
	if ($Installer.PSObject.Properties.Name.Contains('InstallerSuccessCodes'))
	{
		return $Installer.InstallerSuccessCodes
	}
	elseif ($Manifest.PSObject.Properties.Name.Contains('InstallerSuccessCodes'))
	{
		return $Manifest.InstallerSuccessCodes
	}
	else
	{
		# Zero is valid for everything.
		0

		# Factor in two msiexec.exe-specific exit codes.
		if ($FilePath.EndsWith('msi'))
		{
			1641  # Machine needs immediate reboot.
			3010  # Reboot should be rebooted.
		}
	}
}


#---------------------------------------------------------------------------
#
# For `-IgnoreHashFailure`, returns the default arguments to perform a silent install for the given installer.
#
#---------------------------------------------------------------------------

function Get-DefaultKnownSwitches
{
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[System.String]$InstallerType
	)

	# Switch on the installer type and return an array of strings for the args.
	switch -Regex ($InstallerType)
	{
		'^(Burn|Wix|Msi)$'
		{
			"/quiet"
			"/norestart"
			"/log `"$($script.LogFile -f 'WinGet')`""
			break
		}
		'^Nullsoft$'
		{
			"/S"
			break
		}
		'^Inno$'
		{
			"/VERYSILENT"
			"/NORESTART"
			"/LOG=`"$($script.LogFile -f 'WinGet')`""
			break
		}
		default
		{
			throw "The installer type '$_' is unsupported."
		}
	}
}


#---------------------------------------------------------------------------
#
# Gets the installer switches for the piped object and specified type.
#
#---------------------------------------------------------------------------

filter Get-WinGetManifestInstallSwitches
{
	param
	(
		[Parameter(Mandatory = $true, ValueFromPipeline = $true)]
		[ValidateNotNullOrEmpty()]
		[pscustomobject]$InputObject,

		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[System.String]$Type
	)

	# Test whether the piped object has InstallerSwitches and it's not null.
	if (($InputObject.PSObject.Properties.Name -notcontains 'InstallerSwitches') -or ($null -eq $InputObject.InstallerSwitches))
	{
		return
	}

	# Return the requested type. This will be null if its not available.
	return $InputObject.InstallerSwitches.PSObject.Properties | Microsoft.PowerShell.Core\Where-Object {$_.Name -eq $Type} | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty Value
}


#---------------------------------------------------------------------------
#
# For `-IgnoreHashFailure`, returns an argument array for `Start-Process`.
#
#---------------------------------------------------------------------------

function Get-WinGetAppArguments
{
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[pscustomobject]$Manifest,

		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[pscustomobject]$Installer,

		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[System.String]$FilePath
	)

	# Add standard msiexec.exe args.
	if ($FilePath.EndsWith('msi'))
	{
		"/i `"$FilePath`""
	}

	# If we're not overriding, get silent switches from manifest and $Custom if we can.
	if (!$Override)
	{
		# Try to get switches from the installer, then the manifest, then by what the installer is, either from the installer or the manifest.
		if ($switches = $Installer | Get-WinGetManifestInstallSwitches -Type Silent)
		{
			# First check the installer array for a silent switch.
			$switches
			Write-LogEntry -Message "Using Silent switches from the manifest's installer data."
		}
		elseif ($switches = $Manifest | Get-WinGetManifestInstallSwitches -Type Silent)
		{
			# Fall back to the manifest itself.
			$switches
			Write-LogEntry -Message "Using Silent switches from the manifest's top level."
		}
		elseif ($instType = $Installer | Get-WinGetInstallerType)
		{
			# We have no defined switches, try to determine switches from the installer's defined type.
			Get-DefaultKnownSwitches -InstallerType $instType
			Write-LogEntry -Message "Using default switches for the manifest installer's installer type ($instType)."
		}
		elseif ($instType = $Manifest | Get-WinGetInstallerType)
		{
			# The installer array doesn't define a type, see if the manifest itself does.
			Get-DefaultKnownSwitches -InstallerType $instType
			Write-LogEntry -Message "Using default switches for the manifest's installer type ($instType)."
		}
		elseif ($switches = $Installer | Get-WinGetManifestInstallSwitches -Type SilentWithProgress)
		{
			# We're shit out of luck... circle back and see if we have _anything_ we can use.
			$switches
			Write-LogEntry -Message "Using SilentWithProgress switches from the manifest's installer data."
		}
		elseif ($switches = $Manifest | Get-WinGetManifestInstallSwitches -Type SilentWithProgress)
		{
			# Last-ditch effort. It's this or bust.
			$switches
			Write-LogEntry -Message "Using SilentWithProgress switches from the manifest's top level."
		}
		else
		{
			throw "Unable to determine how to silently install the application."
		}

		# Append any custom switches the caller has provided.
		if ($Custom)
		{
			$Custom
		}
	}
	else
	{
		# Override replaces anything the manifest provides.
		$Override
	}
}


#---------------------------------------------------------------------------
#
# Main entry point for script.
#
#---------------------------------------------------------------------------

function Invoke-WinGetAppProcess
{
	# Perform tests prior to commencing operations.
	$systemUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name.Equals('NT AUTHORITY\SYSTEM')
	Test-ConsoleEncoding
	Test-WinGetScope

	# Define variables needed for operations.
	$wgExitCodes = Out-WinGetExitCodeTable
	$wgExecPath = Out-WinGetPath

	# Test whether we're debugging $IgnoreHashFailure.
	if (!($debugging = $script.Action.Equals('install') -and $DebugHashFailure))
	{
		# Set up args for Invoke-WinGetExecutable and commence process.
		$wpParams = @{
			LiteralPath = $wgExecPath
			Arguments = Out-WinGetArgArray
			Silent = $script.Action.Equals('list')
		}
		$wgOutput = Invoke-WinGetExecutable @wpParams

		# If package isn't found, rerun again without --Scope argument.
		if ($LASTEXITCODE.Equals($wgExitCodes.NO_APPLICABLE_INSTALLER))
		{
			Write-LogEntry -Message "Attempting to execute WinGet again without '--scope' argument."
			$wpParams.Arguments = Out-WinGetArgArray -Exclude Scope
			$wgOutput = Invoke-WinGetExecutable @wpParams
		}
	}
	else
	{
		# Going into bypass mode. Simulate WinGet output for the purpose of getting the app's version later on.
		Write-LogEntry -Message "Bypassing WinGet as `-DebugHashFailure` has been passed. This switch should only be used for debugging purposes."
		$wgAppData = & $wgExecPath search --Id $id --exact --accept-source-agreements | Convert-WinGetListOutput
		$wgOutput = [System.String[]]"Found $($wgAppData.Name) [$Id] Version $($wgAppData.Version)."
	}

	# Process resulting exit code.
	if ($debugging -or ($LASTEXITCODE.Equals($wgExitCodes.INSTALLER_HASH_MISMATCH) -and $IgnoreHashFailure))
	{
		# The hash failed, however we're forcing an override.
		Write-LogEntry -Message "Installation failed due to mismatched hash, attempting to override as `-IgnoreHashFailure` has been passed."

		# Munge out the app's version from WinGet's log without having to call WinGet again for it.
		$wgAppVerRegex = "^Found\s.+\s[$([System.Text.RegularExpressions.Regex]::Escape($Id))].+Version\s((\d|\.)+)\.$"
		$wgAppVersion = $($wgOutput -match $wgAppVerRegex -replace $wgAppVerRegex,'$1')

		# Get relevant app information.
		$wgAppInfo = [ordered]@{}
		$wgAppInfo.Add('Manifest', (Get-WinGetAppManifest -AppVersion $wgAppVersion))
		$wgAppInfo.Add('Installer', (Get-WinGetAppInstaller -Manifest $wgAppInfo.Manifest))
		$wgAppInfo.Add('FilePath', (Get-WinGetAppDownload -Installer $wgAppInfo.Installer))

		# Set up arguments to pass to Start-Process.
		$spParams = @{
			WorkingDirectory = $PWD.Path
			ArgumentList = Get-WinGetAppArguments @wgAppInfo
			FilePath = $(if ($wgAppInfo.FilePath.EndsWith('msi')) {'msiexec.exe'} else {$wgAppInfo.FilePath})
			PassThru = $true
			Wait = $true
		}

		# Commence installation and test the resulting exit code for success.
		Write-LogEntry -Message "Starting package install..."
		Write-LogEntry -Message "Executing [$($spParams.FilePath) $($spParams.ArgumentList)]"
		if ((Get-WinGetAppExitCodes @wgAppInfo) -notcontains ($script.ExitCode = (Microsoft.PowerShell.Management\Start-Process @spParams).ExitCode))
		{
			throw "The package installation failed with exit code [$($script.ExitCode)]."
		}

		# Yay, we made it!
		Write-LogEntry -Message "Successfully installed."
	}
	elseif ($wpParams.Silent -and !$LASTEXITCODE)
	{
		# Convert the console output into a proper object.
		$wgAppData = $wgOutput | Convert-WinGetListOutput
		$wgLogBase = "$($wgAppData.Name) [$($wgAppData.Id)] $($wgAppData.Version)"

		# Do some version checking of the found application.
		if (![System.String]::IsNullOrWhiteSpace($Version) -and ([System.Version]($wgAppData.Version -replace '[^\d.]') -lt [System.Version]($Version -replace '[^\d.]')))
		{
			throw "Detected $wgLogBase, but $Version or higher is required."
		}
		elseif ($wgAppData.PSObject.Properties.Name.Contains('Available') -and ([System.Version]($wgAppData.Available -replace '[^\d.]') -gt [System.Version]($wgAppData.Version -replace '[^\d.]')))
		{
			Write-LogEntry -Message "Detected $wgLogBase, but $($wgAppData.Available) is available." -Warning
		}
		else
		{
			Write-LogEntry -Message "Successfully detected $wgLogBase."
		}
	}
	elseif (($script.ExitCode = $LASTEXITCODE))
	{
		# Throw a terminating error message. All this bullshit is to change crap like '0x800704c7 : unknown error.' to 'Unknown error.'...
		$wgErrorDef = $wgExitCodes.PSObject.Properties.Where({$_.Value.Equals($LASTEXITCODE)}) | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty Name
		$wgErrorMsg = [System.Text.RegularExpressions.Regex]::Replace($wgOutput[-1], '^0x\w{8}\s:\s(\w)', {$args[0].Groups[1].Value.ToUpper()})
		throw "WinGet operation finished with exit code 0x$($LASTEXITCODE.ToString('X'))$(if ($wgErrorDef) {" ($wgErrorDef)"}) [$($wgErrorMsg.TrimEnd('.'))]."
	}
}

# Commence function.
if (Open-LogFile)
{
	try
	{
		Invoke-WinGetAppProcess
	}
	catch
	{
		# Handle messages we throw vs. cmdlets throwing exceptions.
		if (($_.CategoryInfo.Category -eq 'OperationStopped') -and ($_.CategoryInfo.Reason -eq 'RuntimeException'))
		{
			# Write the actual error message to StdErr.
			Write-LogEntry -Message $_.Exception.Message -Type Error

			# Advise where WinGet error code definitions can be obtained.
			if (Get-Variable -Name LASTEXITCODE -ValueOnly -ErrorAction Ignore)
			{
				$errCodeUri = 'https://github.com/microsoft/winget-cli/blob/master/doc/windows/package-manager/winget/returnCodes.md'
				Write-LogEntry -Message "Refer to '$errCodeUri' for more information."
			}
		}
		else
		{
			# A native cmdlet has errored out. Pretty up the message before writing it out.
			$_ | Out-FriendlyErrorMessage | Write-LogEntry -Type Error
		}

		# Ensure we exit with some kind of exit code on an error. 1603 is msiexec.exe fatal error.
		if (!$script.ExitCode)
		{
			$script.ExitCode = 1603
		}
	}
	finally
	{
		Close-LogFile
		[System.Console]::OutputEncoding = $lastConsoleOutputEncoding
		exit $script.ExitCode
	}
}
