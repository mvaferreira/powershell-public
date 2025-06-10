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
    Detects ghosted (disconnected) and valid network interface cards (NICs) on Windows.
    Author: Marcus Ferreira marcus.ferreira[at]microsoft[dot]com
    Version: 0.1

.DESCRIPTION
    This script scans the Windows registry for network adapters on PCI and VMBUS buses,
    compares them with currently active network adapters, and identifies ghosted (disconnected)
    NICs as well as valid ones. Useful for troubleshooting network issues or cleaning up old NICs.

.NOTES
    Requires administrator privileges.
    Tested on Windows Server 2016+.

.EXAMPLE
    Run as administrator:
    PS> .\WindowsDetectGhostedNICs.ps1
#>

# Ensure script runs with administrator privileges
If (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script must be run as Administrator."
    exit 1
}

# Get all current network adapter PnP IDs
Write-Host "Collecting all current network adapter PnPDeviceIDs..."
$AllAdaptersPnPIDs = (Get-NetAdapter).PnpDeviceID
$global:FoundGhostNICs = 0
$global:FoundValidNICs = 0

# Scans the registry for NICs under the specified bus and identifies ghosted and valid NICs.
Function ScanRegistryForNICs($BusName) {
    $RootPath = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Enum\"
    $RegistryPath = "$($RootPath.Replace("HKEY_LOCAL_MACHINE\","HKLM:\"))$($BusName)"

    Try {
        If (Test-Path -Path $RegistryPath) {
            Get-ChildItem $RegistryPath -Depth 0 | ForEach-Object {
                $Bus = $_
                $BusPath = $Bus.Name.ToString().Replace("HKEY_LOCAL_MACHINE\","HKLM:\")

                Get-ChildItem $BusPath -Depth 0 | ForEach-Object {
                    $Bus = $_
                    $BusPath = $Bus.Name.ToString().Replace("HKEY_LOCAL_MACHINE\","HKLM:\")
                    $BusRelativeName = $Bus.Name.ToString().Replace($RootPath,"")

                    Try {
                        $ServiceType = (Get-ItemProperty -Path $BusPath -Name Service -ErrorAction SilentlyContinue).Service
                    } Catch {
                        $ServiceType = $null
                    }
                    
                    Try {
                        $DeviceDesc = (Get-ItemProperty -Path $BusPath -Name DeviceDesc -ErrorAction SilentlyContinue).DeviceDesc
                    } Catch {
                        $DeviceDesc = $null
                    }

                    Try {
                        $DevIndex = (Get-ItemProperty -Path $($BusPath + "\Device Parameters") -Name InstanceIndex -ErrorAction SilentlyContinue).InstanceIndex
                    } Catch {
                        $DevIndex = $null
                    }

                    If ($DeviceDesc) {
                        Try {
                            $Split = $DeviceDesc -split ";"
                            $DeviceDescription = $Split[1]

                            If ($DevIndex) {
                                If ($DevIndex -ne 1) {
                                    $DeviceDescription = $DeviceDescription + " #$($DevIndex)"
                                }
                            }
                        } Catch {
                            $DeviceDescription = "Unknown Device"
                        }
                    } Else {
                        $DeviceDescription = "Unknown Device"
                    }

                    If ($ServiceType) {
                        If ($ServiceType -in ("netvsc", "mlx5", "mlx4_bus")) {
                            $MatchesPnpID = $AllAdaptersPnPIDs | Where-Object { $_ -eq $BusRelativeName }

                            If ($MatchesPnpID) {
                                $global:FoundValidNICs++
                                Write-Host "Valid NIC: $($DeviceDescription)" -ForegroundColor Green
                            } Else {
                                $global:FoundGhostNICs++
                                Write-Host "Ghosted NIC: $($DeviceDescription)" -ForegroundColor Yellow
                            }                        
                        }
                    } Else {
                        #Write-Host "Device '$($DeviceDescription)' is not a NIC $($BusPath)."
                    }
                }
            }
        } Else {
            #Write-Host "Path $($RegistryPath) not found, skipping..."
        }
    } Catch {
        Write-Error "Error while scanning registry: $_"
    }
}

ScanRegistryForNICs("PCI")
ScanRegistryForNICs("VMBUS")

Write-Host "`r`n"
Write-Host "Found ghosted NIC(s): $($FoundGhostNICs)" -ForegroundColor Red
Write-Host "Found valid NIC(s): $($FoundValidNICs)" -ForegroundColor Green