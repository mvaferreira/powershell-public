<#
Disclaimer
	The sample scripts are not supported under any Microsoft standard support program or service. 
	The sample scripts are provided AS IS without warranty of any kind.
	Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability
	or of fitness for a particular purpose.
	The entire risk arising out of the use or performance of the sample scripts and documentation remains with you.
	In no event shall Microsoft, its authors, or anyone else involved in the creation, production,
	or delivery of the scripts be liable for any damages whatsoever (including, without limitation,
	damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss)
	arising out of the use of or inability to use the sample scripts or documentation,
	even if Microsoft has been advised of the possibility of such damages.

    .SYNOPSIS
    Author: Marcus Ferreira marcus.ferreira[at]microsoft[dot]com
    Version: 0.1

    .DESCRIPTION
	This script will fix printer server issues by removing third-party Print Processors and Monitors from server.
	This avoids Spooler.exe, PrintIsolationHost.exe and PrintFilterPipeLineSvc.exe from crashing...
	* Tested on Windows Server 2016 *

	Please check url for more details: https://docs.microsoft.com/en-us/archive/blogs/perfguru/print-spooler-crash-troubleshooting-steps
    
    .EXAMPLE
    Run script without parameters to check what the script will do.
	.\FixPrinterServer.ps1
	
	Once done, let the script do its job, with -Modify switch. Make sure you backup the server first...

	.\FixPrinterServer.ps1 -Modify

	This script will backup print root registry key and entire printing with PrintBrm.exe,
	before any changes, but make sure you backup the server first.

	Backups will be available to folder $TempDir (C:\Temp\PrintBackup)
#>

Param(
	[switch]$Modify
)

$TempDir = "C:\Temp\PrintBackup"
$PrintRoot = "HKLM:\SYSTEM\CurrentControlSet\Control\Print"
$PrintEnv = Join-Path $PrintRoot "Environments"
$PrintMonitors = Join-Path $PrintRoot "Monitors"
$PrintMonitorsToKeep = @(
	"Local Port",
	"LPR Port",
	"Standard TCP/IP Port",
	"USB Monitor",
	"WSD Port"
)

#Create temp dir
If (-Not (Test-Path -Path $TempDir -ErrorAction SilentlyContinue)) {
	Write-Host "[$(Get-Date)] Creating backup directory..."
	If ($Modify) { $Null = New-Item -Path $TempDir -ItemType Directory -Force -ErrorAction SilentlyContinue }
}

#Backup registry keys
Write-Host "[$(Get-Date)] Creating Print root key $($PrintRoot) backup..."
If (Test-Path -Path "$TempDir\PrintBackup.reg" -ErrorAction SilentlyContinue) {
	If ($Modify) { $Null = Remove-Item -Confirm:$False "$TempDir\PrintBackup.reg" -ErrorAction SilentlyContinue -Force }
}
If ($Modify) { & $Env:SystemRoot\System32\reg.exe export $PrintRoot.Replace(":", "") "$TempDir\PrintBackup.reg" }

#Backup printer queues, drivers and config via PrintBrm.exe
Write-Host "[$(Get-Date)] Creating Print config backup with PrintBrm.exe..."
If (Test-Path -Path "$TempDir\PrintConfig_$($Env:COMPUTERNAME).printerExport" -ErrorAction SilentlyContinue) {
	If ($Modify) { 
		$Null = Remove-Item -Confirm:$False "$TempDir\PrintConfig_$($Env:COMPUTERNAME).printerExport" -ErrorAction SilentlyContinue -Force
	}
}
If ($Modify) { & $Env:SystemRoot\System32\spool\tools\PrintBrm.exe -B -F "$TempDir\PrintConfig_$($Env:COMPUTERNAME).printerExport" }

#Stop Print Spooler service
Write-Host "[$(Get-Date)] Stopping Spooler service..."
If ($Modify) { Stop-Service -Name "Spooler" -Confirm:$False -Force -ErrorAction SilentlyContinue }

#Kill all PrintIsolationHost.exe and PrintFilterPipeLineSvc.exe processes
Write-Host "[$(Get-Date)] Killing processes 'PrintIsolationHost.exe' and 'PrintFilterPipeLineSvc.exe'..."
If ($Modify) { 
	Stop-Process -Name "PrintIsolationHost" -Confirm:$False -Force -ErrorAction SilentlyContinue
	Stop-Process -Name "PrintFilterPipeLineSvc" -Confirm:$False -Force -ErrorAction SilentlyContinue
}

#Clear spooler directory
$SpoolDirectory = (Get-ItemProperty "$($PrintRoot)\Printers").DefaultSpoolDirectory
Write-Host "[$(Get-Date)] Cleaning Spooler directory $($SpoolDirectory)..."
If ($Modify) { Get-ChildItem -Path $SpoolDirectory -Recurse | Remove-Item -Recurse -Confirm:$False -Force -ErrorAction SilentlyContinue }

#Start Print Spooler service
Write-Host "[$(Get-Date)] Starting Spooler service..."
If ($Modify) { Start-Service -Name "Spooler" -Confirm:$False -ErrorAction SilentlyContinue }

#Set existing print queues to use Print Processor 'winprint'
Get-Printer | ForEach-Object {
	$Printer = $_

	If ($Printer.PrintProcessor -ne "winprint") {
		Write-Host "[$(Get-Date)] Configuring queue $($Printer.Name) print processor $($Printer.PrintProcessor) to 'winprint'"
		If ($Modify) { $Printer | Set-Printer -PrintProcessor "winprint" }
	}
}

#For every environment (x64, x86), set default driver processor and monitor
Get-ChildItem $PrintEnv -Recurse | ForEach-Object {
	$AllEnv = $_

	$AllEnv | Get-ItemProperty | ForEach-Object {
		$Driver = $_

		#Set driver default Print Processor to 'winprint'
		If ($Driver."Print Processor") {
			$DriverPath = $Driver.PSPath.Replace("Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE", "HKLM:")
			If ((Get-ItemProperty $DriverPath)."Print Processor" -ne "winprint") {
				Write-Host "[$(Get-Date)] Changing Print Processor '$((Get-ItemProperty $DriverPath)."PSChildName")' from '$((Get-ItemProperty $DriverPath)."Print Processor")' to 'winprint'"
				If ($Modify) { Set-ItemProperty $DriverPath -Name "Print Processor" -Value "winprint" }
			}
		}

		#Set driver default Print Monitor to 'none/empty'
		If ($Driver."Monitor") {
			$DriverPath = $Driver.PSPath.Replace("Microsoft.PowerShell.Core\Registry::HKEY_LOCAL_MACHINE", "HKLM:")
			If ((Get-ItemProperty $DriverPath)."Monitor" -ne "") {
				Write-Host "[$(Get-Date)] Changing Print Monitor '$((Get-ItemProperty $DriverPath)."PSChildName")' from '$((Get-ItemProperty $DriverPath)."Monitor")' to ' '"
				If ($Modify) { Set-ItemProperty $DriverPath -Name "Monitor" -Value "" }
			}
		}
	}
}

#Delete all third-party Print Processors
Get-ChildItem $PrintEnv | ForEach-Object {
	$Env = $_

	$3rdPartyProcessors = (Join-Path $Env.Name "Print Processors").Replace("HKEY_LOCAL_MACHINE", "HKLM:")
	$3rdPartyProcessors | Get-ChildItem | ForEach-Object {
		$ProcessorKey = $_.ToString().Replace("HKEY_LOCAL_MACHINE", "HKLM:")
		If ($ProcessorKey -notmatch "winprint") {
			Write-Host "[$(Get-Date)] Removing Print Processor $($ProcessorKey)"
			If ($Modify) { Remove-Item $ProcessorKey -Force }
		}
	}
}

#Delete all third-party Print Monitors
Get-ChildItem $PrintMonitors | ForEach-Object {
	$3rdPartyMonitor = $_.Name.Replace("HKEY_LOCAL_MACHINE", "HKLM:")

	If (-Not ($PrintMonitorsToKeep | Where-Object { $_ -eq (Get-Item -Path $3rdPartyMonitor).PSChildName })) {
		Write-Host "[$(Get-Date)] Removing Print Monitor $($3rdPartyMonitor)"
		If ($Modify) { Remove-Item $3rdPartyMonitor -Recurse -Force }
	}
}

#Restart Print Spooler service
Write-Host "[$(Get-Date)] Restarting Spooler service..."
If ($Modify) { Restart-Service -Name "Spooler" -Confirm:$False -Force -ErrorAction SilentlyContinue }