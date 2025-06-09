<#
.SYNOPSIS
    This script identifies and removes "ghosted" Mellanox (VEN_15B3) network interface cards (NICs) from the Windows registry.
.DESCRIPTION
    - Scans the registry for Mellanox NICs.
    - Compares registry entries to currently present network adapters.
    - Identifies ghosted NICs (present in registry but not in the system).
    - Prompts the user to remove ghosted NICs.
    - Removes registry keys for ghosted NICs and cleans up using pnpclean.
    - Handles errors and logs actions for traceability.
.NOTES
    Run as administrator. Some actions require SYSTEM privileges.
#>

# Ensure script runs with administrator privileges
If (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script must be run as Administrator."
    exit 1
}

# Get all current network adapter PnP IDs
Write-Host "Collecting all current network adapter PnPDeviceIDs..."
$AllAdaptersPnPIDs = (Get-NetAdapter).PnpDeviceID
$FoundGhostNICs = 0
$FoundValidNICs = 0
$GhostedNICsToDelete = @()

#Executes a PowerShell command as the SYSTEM account using a temporary scheduled task. 
#This function creates a scheduled task that runs the specified command as SYSTEM.
Function ExecuteAsSystem($cmd) {
    #Write-Host "Executing command as SYSTEM: $cmd"
    Try {
        $taskName = "TempSystemTask_$([guid]::NewGuid().ToString())"
        $TaskArgs = '-NoProfile -WindowStyle Hidden -Command "###cmd###"'
        $TaskArgs = $TaskArgs -replace '###cmd###',$cmd
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $TaskArgs
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1)

        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal | Out-Null

        Start-ScheduledTask -TaskName $taskName

        While ((Get-ScheduledTask -TaskName $taskName).State -eq "Running" -Or (Get-ScheduledTaskInfo -TaskName $taskName).LastRunTime -notmatch (Get-Date).ToString("yyyy")) {
            Start-Sleep -Seconds 1
        }

        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false | Out-Null
    } Catch {
        Write-Error "Failed to execute command as SYSTEM: $_"
    }
}

# Checks for empty Mellanox PCI root registry keys and deletes them if found.
# Scans the PCI registry for Mellanox (VEN_15B3) keys with no subkeys and deletes them using ExecuteAsSystem.
Function CheckRootKeys() {
    Write-Host "Checking for empty Mellanox PCI root keys in registry..."
    $GhostedNICsToDelete = @()

    Try {
        Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Enum\PCI" -Depth 0 | ForEach-Object {
            $Bus = $_
            If ($Bus.Name -match "VEN_15B3") {
                $BusPathMlx = $Bus.Name.ToString().Replace("HKEY_LOCAL_MACHINE\","HKLM:\")

                # Mark root key to be deleted if it has no subkeys
                If ((Get-ChildItem $BusPathMlx -Depth 0 | Measure-Object).Count -eq 0) {
                    $GhostedNICsToDelete += $BusPathMlx
                }
            }
        }
    } Catch {
        Write-Error "Error while checking root keys: $_"
    }

    $GhostedNICsToDelete | ForEach-Object {
        If (Test-Path -Path $_) {
            Write-Host "Deleting empty root key $($_)..."
            ExecuteAsSystem "Remove-Item -Path '$($_)' -Recurse -Confirm:`$false -Force"
        }
    }
}

# Removes ghosted Mellanox NIC registry keys and runs pnpclean.
# Invokes pnpclean to clean up device references, then deletes registry keys for ghosted NICs using ExecuteAsSystem.
Function DeleteGhostedNICs() {
    Write-Host "Clearing ghosted NIC(s) using pnpclean..."
    Try {
        Invoke-Command {C:\windows\system32\RUNDLL32.exe c:\windows\system32\pnpclean.dll,RunDLL_PnpClean /Devices /Maxclean}
        Start-Sleep -Seconds 10

        $GhostedNICsToDelete | ForEach-Object {
            If (Test-Path -Path $_) {
                Write-Host "Deleting registry key for ghosted NIC: $($_)..."
                ExecuteAsSystem "Remove-Item -Path '$($_)' -Recurse -Confirm:`$false -Force"
            }
        }
    } Catch {
        Write-Error "Error while deleting ghosted NICs: $_"
    }

    CheckRootKeys
}

# Main logic: Scans registry for Mellanox NICs, compares with present adapters, and prompts for cleanup.
# Enumerates Mellanox NICs in the registry, checks if they are present in the system, and identifies ghosted NICs.
# Scan registry for Mellanox NICs and compare with current adapters
Write-Host "Scanning registry for Mellanox (VEN_15B3) NICs..."
Write-Host "`r`n"
Try {
    Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Enum\PCI" -Depth 0 | ForEach-Object {
        $Bus = $_
        If ($Bus.Name -match "VEN_15B3") {
            $BusPathMlx = $Bus.Name.ToString().Replace("HKEY_LOCAL_MACHINE\","HKLM:\")
            Get-ChildItem $BusPathMlx | ForEach-Object {
                $Bus = $_
                $BusPath = $Bus.Name.ToString().Replace("HKEY_LOCAL_MACHINE\","HKLM:\")
                $BusRelativeName = $Bus.Name.ToString().Replace("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Enum\","")
                
                Try {
                    $DeviceDesc = (Get-ItemProperty -Path $BusPath -Name DeviceDesc).DeviceDesc
                } Catch {
                    $DeviceDesc = $null
                }

                Try {
                    $DevIndex = (Get-ItemProperty -Path $($BusPath + "\Device Parameters") -Name InstanceIndex -ErrorAction SilentlyContinue).InstanceIndex
                } Catch {
                    $DevIndex = $null
                }

                If ($DeviceDesc) {
                    $Split = $DeviceDesc -split ";"
                    $DeviceDescription = $Split[1]

                    If ($DevIndex) {
                        If ($DevIndex -ne 1) {
                            $DeviceDescription = $DeviceDescription + " #$($DevIndex)"
                        }
                    }
                } Else {
                    $DeviceDescription = "Unknown Device"
                }

                $MatchesPnpID = $AllAdaptersPnPIDs | Where-Object { $_ -eq $BusRelativeName }

                If ($MatchesPnpID) {
                    $FoundValidNICs++
                    Write-Host "Valid NIC: $($DeviceDescription)"
                } Else {
                    $FoundGhostNICs++
                    Write-Host "Ghosted NIC: $($DeviceDescription)"
                    $GhostedNICsToDelete += $BusPath
                }
            }
        } Else {
            # Not a Mellanox NIC, ignoring.
        }
    }
} Catch {
    Write-Error "Error while scanning registry: $_"
}

Write-Host "`r`n"
Write-Host "Found ghosted NIC(s): $($FoundGhostNICs)"
Write-Host "Found valid NIC(s): $($FoundValidNICs)"

If ($FoundGhostNICs -gt 0) {
    Write-Host "`r`n"
    $ClearEntries = Read-Host "Should we clear ghosted NIC(s)? (Y/N)"
    Write-Host "`r`n"

    Switch($ClearEntries.ToLower()) {
        { $_ -in "y","yes" } {
            DeleteGhostedNICs
        }
        { $_ -in "n","no" } { 
            Write-Host "No changes made."
        }
        default { 
            Write-Host "Invalid answer. (Y/N)"
            return
        }
    }
}