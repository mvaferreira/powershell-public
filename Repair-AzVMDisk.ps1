<#
    Disclaimer
        The sample scripts are not supported under any Microsoft standard support program or service.
        The sample scripts are provided AS IS without warranty of any kind.
        Microsoft further disclaims all implied warranties including, without limitation, any implied
        warranties of merchantability or of fitness for a particular purpose.
        The entire risk arising out of the use or performance of the sample scripts and documentation
        remains with you. In no event shall Microsoft, its authors, or anyone else involved in the
        creation, production, or delivery of the scripts be liable for any damages whatsoever
        (including, without limitation, damages for loss of business profits, business interruption,
        loss of business information, or other pecuniary loss) arising out of the use of or inability
        to use the sample scripts or documentation, even if Microsoft has been advised of the
        possibility of such damages.

    .SYNOPSIS
        Offline Azure VM disk repair and diagnostic script for use on a Hyper-V rescue VM.
        Author: Marcus Ferreira marcus.ferreira[at]microsoft[dot]com
        Version: 0.2

    .DESCRIPTION
        Repair-AzVMDisk.ps1 attaches the OS disk of a broken Azure VM to a Hyper-V rescue VM and performs
        offline repairs without booting the guest. It can mount offline registry hives
        (BROKENSYSTEM / BROKENSOFTWARE), run chkdsk/SFC/DISM, rebuild BCD, fix RDP/NLA settings,
        manage drivers and services, reset credentials, and collect diagnostic information.

        A built-in system check (-SysCheck) inspects the offline disk for common issues across
        disk health, boot configuration, RDP/NLA policy, Windows Update/CBS state, credential guard,
        network bindings, Azure VM agent presence, and more — and prints actionable fix suggestions.

        All actions are logged to a JSON-line audit file (Repair-AzVMDisk_actions.log) alongside the script.
        Previous sessions can be reviewed with -ShowLastSession.

    .NOTES
        - Must be run as Administrator on the Hyper-V rescue VM.
        - The broken VM's OS disk must be attached to the rescue VM.
        - Tested on Windows Server 2016 / 2019 / 2022 rescue environments.
        - Use -LeaveDiskOnline to keep the disk online after repair (required for chkdsk via -FixNTFS).

    .PARAMETER DiskNumber
        Disk number of the attached broken VM disk (from Get-Disk). Use instead of -VMName.

    .PARAMETER VMName
        Name of the Hyper-V VM whose disk should be attached automatically. Use instead of -DiskNumber.

    .EXAMPLE
        # Run a full diagnostic check on disk 3
        PS> .\Repair-AzVMDisk.ps1 -DiskNumber 3 -SysCheck

    .EXAMPLE
        # Rebuild BCD and fix RDP settings on disk 3
        PS> .\Repair-AzVMDisk.ps1 -DiskNumber 3 -FixBoot -FixRDP

    .EXAMPLE
        # Run chkdsk on a specific partition (disk must stay online for the drive letter to be available)
        PS> .\Repair-AzVMDisk.ps1 -DiskNumber 3 -FixNTFS -DriveLetter H: -LeaveDiskOnline

    .EXAMPLE
        # Reset a forgotten local admin password
        PS> .\Repair-AzVMDisk.ps1 -DiskNumber 3 -ResetLocalAdminPassword

    .EXAMPLE
        # Disable a problematic third-party driver (or several)
        PS> .\Repair-AzVMDisk.ps1 -DiskNumber 3 -DisableDriver driver1,driver2

    .EXAMPLE
        # Show the last repair session log
        PS> .\Repair-AzVMDisk.ps1 -ShowLastSession -Detailed

    .EXAMPLE
        # Export all sessions to HTML
        PS> .\Repair-AzVMDisk.ps1 -ShowLastSession -All -ExportTo C:\Temp\repair_log.html
#>
[CmdletBinding(DefaultParameterSetName = 'Repair')]
param (
    [Parameter(ParameterSetName = 'Repair')][string]$VMName = "",
    [Parameter(ParameterSetName = 'Repair')][int]$DiskNumber = -1,
    [Parameter(ParameterSetName = 'Repair')][switch]$FixNTFS,
    [Parameter(ParameterSetName = 'Repair')][string]$DriveLetter = '',
    [Parameter(ParameterSetName = 'Repair')][switch]$FixBoot,
    [Parameter(ParameterSetName = 'Repair')][switch]$FixBootSector,
    [Parameter(ParameterSetName = 'Repair')][switch]$FixHealth,
    [Parameter(ParameterSetName = 'Repair')][switch]$TryLGKC,
    [Parameter(ParameterSetName = 'Repair')][switch]$TryOtherBootConfig,
    [Parameter(ParameterSetName = 'Repair')][switch]$TrySafeMode,
    [Parameter(ParameterSetName = 'Repair')][switch]$RemoveSafeModeFlag,
    [Parameter(ParameterSetName = 'Repair')][switch]$RunSFC,
    [Parameter(ParameterSetName = 'Repair')][switch]$EnableBootLog,
    [Parameter(ParameterSetName = 'Repair')][switch]$DisableStartupRepair,
    [Parameter(ParameterSetName = 'Repair')][switch]$EnableStartupRepair,
    [Parameter(ParameterSetName = 'Repair')][switch]$DisableBFE,
    [Parameter(ParameterSetName = 'Repair')][switch]$EnableBFE,
    [Parameter(ParameterSetName = 'Repair')][switch]$AddTempUser,
    [Parameter(ParameterSetName = 'Repair')][switch]$AddTempUser2,
    [Parameter(ParameterSetName = 'Repair')][switch]$ResetLocalAdminPassword,
    [Parameter(ParameterSetName = 'Repair')][switch]$SetFullMemDump,
    [Parameter(ParameterSetName = 'Repair')][switch]$DisableNLA,
    [Parameter(ParameterSetName = 'Repair')][switch]$EnableNLA,
    [Parameter(ParameterSetName = 'Repair')][switch]$EnableWinRMHTTPS,
    [Parameter(ParameterSetName = 'Repair')][switch]$CheckRDPPolicies,
    [Parameter(ParameterSetName = 'Repair')][switch]$FixRDP,
    [Parameter(ParameterSetName = 'Repair')][switch]$FixRDPCert,
    [Parameter(ParameterSetName = 'Repair')][switch]$FixRDPPermissions,
    [Parameter(ParameterSetName = 'Repair')][switch]$FixRDPAuth,
    [Parameter(ParameterSetName = 'Repair')][switch]$FixUserRights,
    [Parameter(ParameterSetName = 'Repair')][switch]$FixPendingUpdates,
    [Parameter(ParameterSetName = 'Repair')][switch]$DisableWindowsUpdate,
    [Parameter(ParameterSetName = 'Repair')][switch]$RestoreRegistryFromRegBack,
    [Parameter(ParameterSetName = 'Repair')][switch]$EnableRegBackup,
    [Parameter(ParameterSetName = 'Repair')][switch]$DisableThirdPartyDrivers,
    [Parameter(ParameterSetName = 'Repair')][switch]$EnableThirdPartyDrivers,
    [Parameter(ParameterSetName = 'Repair')][switch]$GetServicesReport,
    [Parameter(ParameterSetName = 'Repair')][switch]$IncludeServices,
    [Parameter(ParameterSetName = 'Repair')][switch]$IssuesOnly,
    [Parameter(ParameterSetName = 'Repair')][string[]]$DisableDriver = @(),
    [Parameter(ParameterSetName = 'Repair')][string[]]$EnableDriver  = @(),
    [Parameter(ParameterSetName = 'Repair')][ValidateSet('Boot','System','Automatic','Manual','Disabled')][string]$DriverStartType = 'Manual',
    [Parameter(ParameterSetName = 'Repair')][switch]$DisableCredentialGuard,
    [Parameter(ParameterSetName = 'Repair')][switch]$EnableCredentialGuard,
    [Parameter(ParameterSetName = 'Repair')][switch]$DisableAppLocker,
    [Parameter(ParameterSetName = 'Repair')][switch]$FixSanPolicy,
    [Parameter(ParameterSetName = 'Repair')][switch]$FixAzureGuestAgent,
    [Parameter(ParameterSetName = 'Repair')][switch]$InstallAzureVMAgent,
    [Parameter(ParameterSetName = 'Repair')][switch]$FixDeviceFilters,
    [Parameter(ParameterSetName = 'Repair')][switch]$KeepDefaultFilters,
    [Parameter(ParameterSetName = 'Repair')][switch]$CopyACPISettings,
    [Parameter(ParameterSetName = 'Repair')][switch]$ScanNetBindings,
    [Parameter(ParameterSetName = 'Repair')][switch]$FixNetBindings,
    [Parameter(ParameterSetName = 'Repair')][switch]$SysCheck,
    [Parameter(ParameterSetName = 'Repair')][switch]$ResetNetworkStack,
    [Parameter(ParameterSetName = 'Repair')][switch]$EnableTestSigning,
    [Parameter(ParameterSetName = 'Repair')][switch]$DisableTestSigning,
    [Parameter(ParameterSetName = 'Repair')][switch]$CheckDiskHealth,
    [Parameter(ParameterSetName = 'Repair')][switch]$CollectEventLogs,
    [Parameter(ParameterSetName = 'Repair')][switch]$LeaveDiskOnline,
    [Parameter(ParameterSetName = 'Repair')][ValidateSet('SYSTEM','SOFTWARE','COMPONENTS','SAM','SECURITY')][string[]]$LoadHive  = @(),
    [Parameter(ParameterSetName = 'Repair')][ValidateSet('SYSTEM','SOFTWARE','COMPONENTS','SAM','SECURITY')][string[]]$UnloadHive = @(),
    [Parameter(ParameterSetName = 'ShowSession')][switch]$ShowLastSession,
    [Parameter(ParameterSetName = 'ShowSession')][switch]$Detailed,
    [Parameter(ParameterSetName = 'ShowSession')][switch]$All,
    [Parameter(ParameterSetName = 'ShowSession')][string]$SessionId = '',
    [Parameter(ParameterSetName = 'ShowSession')][string]$ExportTo = ''
)

################################################################################
# Helper functions
################################################################################

# Helper: builds a unique backup file path if the primary .bak exists
function New-UniqueBackupPath {
    param(
        [Parameter(Mandatory = $true)][string]$BasePath, # e.g. E:\Boot\BCD or H:\EFI\Microsoft\Boot\BCD
        [string]$BakSuffix = ".bak"
    )
    $bak = "$BasePath$BakSuffix"
    if (-not (Test-Path -LiteralPath $bak)) { return $bak }
    $ts = (Get-Date -Format "yyyyMMddHHmmss")
    return "$BasePath.$ts$BakSuffix"
}

# Helper: Get BCD store path based on VM generation
function Get-BcdStorePath {
    param(
        [Parameter(Mandatory = $true)][int]$Generation,
        [Parameter(Mandatory = $true)][string]$BootDrive
    )
    $BootDrive = $BootDrive.TrimEnd('\')
    if ($Generation -eq 1) {
        return "$BootDrive\Boot\BCD"
    }
    else {
        return "$BootDrive\EFI\Microsoft\Boot\BCD"
    }
}

# Helper: Extract Windows Boot Loader identifier from BCD
function Get-BcdBootLoaderId {
    param(
        [Parameter(Mandatory = $true)][string]$StorePath
    )
    if (-not (Test-Path -LiteralPath $StorePath)) {
        Write-Warning "BCD not found at $StorePath."
        return $null
    }
    
    $cmd = "bcdedit /store `"$StorePath`" /enum"
    $raw = & cmd.exe /c $cmd
    $identifier = (($raw -join "`n") -replace '(?s).*Windows Boot Loader.*?identifier\s+([^\r\n]+).*', '$1').Trim()
    
    if ([string]::IsNullOrWhiteSpace($identifier)) {
        Write-Warning "Could not extract boot loader identifier from $StorePath"
        return $null
    }
    return $identifier
}

# Helper: Execute bcdedit command with validation
function Invoke-BcdEdit {
    param(
        [Parameter(Mandatory = $true)][string]$StorePath,
        [Parameter(Mandatory = $true)][string]$Command
    )
    $fullCmd = "bcdedit /store `"$StorePath`" $Command"
    Write-Host "  [exec] $fullCmd" -ForegroundColor DarkGray
    Invoke-Logged -Description "bcdedit" -Details @{ StorePath = $StorePath; Command = $Command } -ScriptBlock { & cmd.exe /c $fullCmd }
}

################################################################################
# Action logging helpers
################################################################################

# Path to the action log (JSON line entries)
if (-not $env:RepairActionLog) {
    $script:ActionLogPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) 'Repair-AzVMDisk_actions.log'
} else {
    # Validate that the env override is a simple file path (no UNC, no alternate data streams)
    $envPath = $env:RepairActionLog
    if ($envPath -match '^\\\\' -or $envPath -match ':.*:') {
        Write-Warning "RepairActionLog path '$envPath' looks unsafe; using default log location."
        $script:ActionLogPath = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) 'Repair-AzVMDisk_actions.log'
    } else {
        $script:ActionLogPath = $envPath
    }
}

# Restrict the log file to Administrators and SYSTEM only (no other local users)
if (-not (Test-Path -LiteralPath $script:ActionLogPath)) {
    $null = New-Item -Path $script:ActionLogPath -ItemType File -Force
}
try {
    $acl = Get-Acl -LiteralPath $script:ActionLogPath
    $acl.SetAccessRuleProtection($true, $false)   # disable inheritance, remove inherited entries
    $acl.Access | ForEach-Object { $acl.RemoveAccessRule($_) | Out-Null }
    $adminRule  = [System.Security.AccessControl.FileSystemAccessRule]::new('BUILTIN\Administrators', 'FullControl', 'Allow')
    $systemRule = [System.Security.AccessControl.FileSystemAccessRule]::new('NT AUTHORITY\SYSTEM',    'FullControl', 'Allow')
    $acl.AddAccessRule($adminRule)
    $acl.AddAccessRule($systemRule)
    Set-Acl -LiteralPath $script:ActionLogPath -AclObject $acl
} catch {
    Write-Warning "Could not restrict log file permissions: $_"
}

# Unique identifier for this script execution – stamped on every log entry
$script:CurrentSessionId = [guid]::NewGuid().ToString()

function Start-ActionLog {
    param([string]$HeaderMessage = "Repair actions log")
    $entry = @{
        SessionId  = $script:CurrentSessionId
        Time       = (Get-Date).ToString('o')
        Event      = 'SessionStart'
        Message    = $HeaderMessage
        Parameters = ($PSBoundParameters.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ' '
    }
    $entry | ConvertTo-Json -Depth 5 -Compress | Out-File -FilePath $script:ActionLogPath -Encoding UTF8 -Append
}

function Write-ActionLog {
    param(
        [Parameter(Mandatory = $true)][string]$Event,
        [Parameter(Mandatory = $true)][hashtable]$Details
    )
    $payload = @{
        SessionId = $script:CurrentSessionId
        Time      = (Get-Date).ToString('o')
        Event     = $Event
        Details   = $Details
    }
    $payload | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath $script:ActionLogPath -Encoding UTF8 -Append
}

# Execute a scriptblock and log start/end, success/failure and outputs
function Invoke-Logged {
    param(
        [Parameter(Mandatory = $true)][string]$Description,
        [hashtable]$Details = @{},
        [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock
    )

    $start = Get-Date
    try {
        $output = & $ScriptBlock 2>&1 | Out-String
        $end = Get-Date
        Write-ActionLog -Event 'ActionExecuted' -Details @{ Description = $Description; Start = $start.ToString('o'); End = $end.ToString('o'); Success = $true; Details = $Details; Output = $output }
        return $output
    }
    catch {
        $end = Get-Date
        Write-ActionLog -Event 'ActionExecuted' -Details @{ Description = $Description; Start = $start.ToString('o'); End = $end.ToString('o'); Success = $false; Details = $Details; Error = $_.ToString() }
        throw
    }
}

# Display all log entries from the most recent session (or a specific session by GUID)
function Get-LastRepairSession {
    param(
        [string]$LogPath = $script:ActionLogPath,
        [string]$SessionId = '',        # leave blank for the most recent session
        [switch]$All,                   # show all sessions
        [switch]$Detailed,              # show full details (paths, output, registry changes, etc.)
        [string]$ExportTo = ''          # optional path to export HTML (.html) or CSV (.csv)
    )

    if (-not (Test-Path $LogPath)) {
        Write-Warning "Log file not found: $LogPath"
        return
    }

    $entries = Get-Content $LogPath | ForEach-Object {
        try { $_ | ConvertFrom-Json } catch { $null }
    } | Where-Object { $_ }

    if ($All) {
        # When -All and -ExportTo are combined, export all entries; otherwise print them
        if ($ExportTo) {
            # fall through: $entries already contains everything, export block below handles it
        } else {
            $entries | Format-List
            return
        }
    } elseif ($SessionId) {
        $entries = $entries | Where-Object { $_.SessionId -eq $SessionId }
    }
    else {
        # Pick the most recent SessionStart entry and filter by its GUID
        $lastSession = ($entries | Where-Object { $_.Event -eq 'SessionStart' } | Select-Object -Last 1).SessionId
        if (-not $lastSession) {
            Write-Warning "No session start marker found in log."
            return
        }
        Write-Host "Showing session: $lastSession" -ForegroundColor Cyan
        $entries = $entries | Where-Object { $_.SessionId -eq $lastSession }
    }

    # Print to console (skip when -All -ExportTo is used — too verbose for hundreds of entries)
    if (-not ($All -and $ExportTo)) {
        foreach ($e in $entries) {
            $entryTime = $e.Time
            $entryType = $e.Event
            $success   = if ($null -ne $e.Details.Success) { if ($e.Details.Success) { '[OK]' } else { '[FAIL]' } } else { '' }
            $desc      = if ($e.Details.Description) { $e.Details.Description } elseif ($e.Message) { $e.Message } else { '' }
            $color     = if ($e.Details.Success -eq $false) { 'Red' } elseif ($entryType -eq 'SessionStart') { 'Cyan' } else { 'White' }
            Write-Host "$entryTime  $entryType  $success  $desc" -ForegroundColor $color
            if ($e.Details.Error) { Write-Host "  ERROR: $($e.Details.Error)" -ForegroundColor Red }

            if ($Detailed -and $entryType -ne 'SessionStart') {
                # Duration
                if ($e.Details.Start -and $e.Details.End) {
                    $duration = ([datetime]$e.Details.End - [datetime]$e.Details.Start).TotalSeconds
                    Write-Host ("    Duration: {0:N3}s" -f $duration) -ForegroundColor DarkGray
                }
                # Operation-specific fields stored in the nested Details sub-object
                if ($e.Details.Details) {
                    $e.Details.Details.PSObject.Properties | Where-Object { $null -ne $_.Value -and $_.Value -ne '' } | ForEach-Object {
                        Write-Host "    $($_.Name): $($_.Value)" -ForegroundColor DarkCyan
                    }
                }
                # Command / tool output
                if ($e.Details.Output -and ([string]$e.Details.Output).Trim()) {
                    Write-Host "    Output:" -ForegroundColor DarkGray
                    ([string]$e.Details.Output).Trim() -split "`n" | ForEach-Object { Write-Host "      $_" -ForegroundColor Gray }
                }
            }
        } # end foreach
    } # end if not (All+ExportTo)

    # --- Export to file if requested ---
    if ($ExportTo) {
        $ext = [System.IO.Path]::GetExtension($ExportTo).ToLower()

        if ($ext -eq '.csv') {
            $rows = foreach ($e in $entries) {
                $duration = ''
                if ($e.Details.Start -and $e.Details.End) {
                    $duration = '{0:N3}s' -f ([datetime]$e.Details.End - [datetime]$e.Details.Start).TotalSeconds
                }
                $detailFields = if ($e.Details.Details) {
                    ($e.Details.Details.PSObject.Properties | Where-Object { $null -ne $_.Value -and $_.Value -ne '' } | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join '; '
                } else { '' }
                [PSCustomObject]@{
                    Time        = $e.Time
                    Event       = $e.Event
                    Success     = if ($null -ne $e.Details.Success) { $e.Details.Success } else { '' }
                    Description = if ($e.Details.Description) { $e.Details.Description } elseif ($e.Message) { $e.Message } else { '' }
                    Duration    = $duration
                    Details     = $detailFields
                    Error       = if ($e.Details.Error) { $e.Details.Error } else { '' }
                    Output      = if ($e.Details.Output) { ([string]$e.Details.Output).Trim() } else { '' }
                }
            }
            $rows | Export-Csv -Path $ExportTo -NoTypeInformation -Encoding UTF8
            Write-Host "Session exported to CSV: $ExportTo" -ForegroundColor Green

        } elseif ($ext -eq '.html') {
            $htmlRows = foreach ($e in $entries) {
                $entryType = $e.Event
                $success   = if ($null -ne $e.Details.Success) { if ($e.Details.Success) { '<span style="color:green">[OK]</span>' } else { '<span style="color:red">[FAIL]</span>' } } else { '' }
                $desc      = [System.Web.HttpUtility]::HtmlEncode($(if ($e.Details.Description) { $e.Details.Description } elseif ($e.Message) { $e.Message } else { '' }))
                $rowStyle  = if ($e.Details.Success -eq $false) { 'background:#ffe0e0' } elseif ($entryType -eq 'SessionStart') { 'background:#e0f0ff' } else { '' }
                $duration  = ''
                if ($e.Details.Start -and $e.Details.End) {
                    $duration = '{0:N3}s' -f ([datetime]$e.Details.End - [datetime]$e.Details.Start).TotalSeconds
                }
                $detailHtml = ''
                if ($e.Details.Details) {
                    $detailHtml = '<ul style="margin:2px 0;padding-left:16px;font-size:0.85em;color:#555">' +
                        (($e.Details.Details.PSObject.Properties | Where-Object { $null -ne $_.Value -and $_.Value -ne '' } | ForEach-Object {
                            "<li><b>$([System.Web.HttpUtility]::HtmlEncode($_.Name)):</b> $([System.Web.HttpUtility]::HtmlEncode([string]$_.Value))</li>"
                        }) -join '') + '</ul>'
                }
                $outputHtml = ''
                if ($e.Details.Output -and ([string]$e.Details.Output).Trim()) {
                    $outputHtml = '<pre style="margin:2px 0;font-size:0.8em;background:#f5f5f5;padding:4px;white-space:pre-wrap">' +
                        [System.Web.HttpUtility]::HtmlEncode(([string]$e.Details.Output).Trim()) + '</pre>'
                }
                $errorHtml = ''
                if ($e.Details.Error) {
                    $errorHtml = '<div style="color:red;font-size:0.85em">ERROR: ' + [System.Web.HttpUtility]::HtmlEncode($e.Details.Error) + '</div>'
                }
                "<tr style='$rowStyle'><td style='white-space:nowrap'>$($e.Time)</td><td>$entryType</td><td>$success</td><td>$desc</td><td>$duration</td><td>$detailHtml$errorHtml$outputHtml</td></tr>"
            }
            $sessionLabel = if ($All) { 'AllSessions' } elseif ($SessionId) { $SessionId } else { ($entries | Where-Object { $_.Event -eq 'SessionStart' } | Select-Object -First 1).SessionId }
            $html = @"
<!DOCTYPE html>
<html><head><meta charset='utf-8'>
<title>Repair-AzVMDisk Session $sessionLabel</title>
<style>
  body { font-family: Segoe UI, Arial, sans-serif; font-size: 13px; margin: 20px; }
  h2 { color: #1a5276; }
  table { border-collapse: collapse; width: 100%; }
  th { background: #1a5276; color: white; padding: 6px 10px; text-align: left; }
  td { border: 1px solid #ddd; padding: 5px 10px; vertical-align: top; }
  tr:hover { background: #f0f7ff; }
</style>
</head><body>
<h2>Repair-AzVMDisk Session Report</h2>
<p><b>Session ID:</b> $sessionLabel<br><b>Log:</b> $LogPath<br><b>Generated:</b> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
<table>
<tr><th>Time</th><th>Event</th><th>Result</th><th>Description</th><th>Duration</th><th>Details / Output</th></tr>
$($htmlRows -join "`n")
</table></body></html>
"@
            $html | Out-File -FilePath $ExportTo -Encoding UTF8
            Write-Host "Session exported to HTML: $ExportTo" -ForegroundColor Green

        } else {
            Write-Warning "Unknown export format '$ext'. Use .html or .csv"
        }
    }
}

# File operation wrappers
function Move-Item-Logged {
    param([Parameter(Mandatory = $true)]$LiteralPath, [Parameter(Mandatory = $true)]$Destination, [switch]$Force)
    $details = @{ Operation = 'Move-Item'; Path = $LiteralPath; Destination = $Destination }
    Invoke-Logged -Description 'Move-Item' -Details $details -ScriptBlock { Move-Item -LiteralPath $LiteralPath -Destination $Destination -Force:$Force }
}

function Remove-Item-Logged {
    param([Parameter(Mandatory = $true)]$Path, [switch]$Recurse, [switch]$Force)
    $recurseStr = if ($Recurse) { ' -Recurse' } else { '' }
    Write-Host "  [exec] Remove-Item$recurseStr '$Path'" -ForegroundColor DarkGray
    $details = @{ Operation = 'Remove-Item'; Path = $Path }
    Invoke-Logged -Description 'Remove-Item' -Details $details -ScriptBlock { Remove-Item -Path $Path -Recurse:$Recurse -Force:$Force -ErrorAction SilentlyContinue }
}

function Rename-Item-Logged {
    param([Parameter(Mandatory = $true)]$Path, [Parameter(Mandatory = $true)]$NewName)
    Write-Host "  [exec] Rename-Item '$Path' -> '$NewName'" -ForegroundColor DarkGray
    $details = @{ Operation = 'Rename-Item'; Path = $Path; NewName = $NewName }
    Invoke-Logged -Description 'Rename-Item' -Details $details -ScriptBlock { Rename-Item -Path $Path -NewName $NewName -Force }
}

# Registry helper that logs before/after values
function Set-ItemProperty-Logged {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)]$Value,
        [Alias('Type')][string]$PropertyType = 'DWord',
        [switch]$Force
    )
    $before = $null
    try { $before = (Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue).$Name } catch { $before = $null }
    $displayValue = if ($Value -is [byte[]]) { "0x$(($Value | ForEach-Object { $_.ToString('X2') }) -join '')" } else { $Value }
    $displayBefore = if ($null -eq $before) { '(not set)' } elseif ($before -is [byte[]]) { "0x$(($before | ForEach-Object { $_.ToString('X2') }) -join '')" } else { $before }
    Write-Host "  [exec] Set-ItemProperty '$Path' -Name '$Name' -Value $displayValue ($PropertyType)  [was: $displayBefore]" -ForegroundColor DarkGray
    $details = @{ Operation = 'Set-ItemProperty'; Path = $Path; Name = $Name; Before = $before; After = $Value }
    Invoke-Logged -Description 'Set-ItemProperty' -Details $details -ScriptBlock { if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }; New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $PropertyType -Force | Out-Null }
}

# Create item wrapper (files/directories/registry keys)
function New-Item-Logged {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Name = '',
        [string]$ItemType = '',
        $Value = $null,
        [switch]$Force,
        [switch]$RedactValue   # when set, logs '***REDACTED***' instead of the actual value
    )
    $logValue = if ($RedactValue) { '***REDACTED***' } else { $Value }
    $displayParts = @("'$Path'")
    if ($Name)     { $displayParts += "-Name '$Name'" }
    if ($ItemType) { $displayParts += "-ItemType '$ItemType'" }
    if ($null -ne $logValue -and $logValue -ne '') { $displayParts += "-Value '$logValue'" }
    Write-Host "  [exec] New-Item $($displayParts -join ' ')" -ForegroundColor DarkGray
    $details = @{ Operation = 'New-Item'; Path = $Path; Name = $Name; ItemType = $ItemType; Value = $logValue }
    Invoke-Logged -Description 'New-Item' -Details $details -ScriptBlock {
        if ($Name -and $ItemType -and $null -ne $Value) {
            New-Item -Path $Path -Name $Name -ItemType $ItemType -Value $Value -Force:$Force | Out-Null
        }
        elseif ($Name -and $ItemType) {
            New-Item -Path $Path -Name $Name -ItemType $ItemType -Force:$Force | Out-Null
        }
        elseif ($ItemType) {
            New-Item -Path $Path -ItemType $ItemType -Force:$Force | Out-Null
        }
        else {
            New-Item -Path $Path -Force:$Force | Out-Null
        }
    }
}

# Remove-ItemProperty wrapper (registry)
function Remove-ItemProperty-Logged {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Name,
        [switch]$Force
    )
    Write-Host "  [exec] Remove-ItemProperty '$Path' -Name '$Name'" -ForegroundColor DarkGray
    $details = @{ Operation = 'Remove-ItemProperty'; Path = $Path; Name = $Name }
    Invoke-Logged -Description 'Remove-ItemProperty' -Details $details -ScriptBlock { Remove-ItemProperty -Path $Path -Name $Name -Force:$Force }
}

# Copy (file/directory) wrapper
function Copy-Item-Logged {
    param(
        [Parameter(Mandatory = $true)]$Path,
        [Parameter(Mandatory = $true)]$Destination,
        [switch]$Recurse,
        [switch]$Force
    )
    $recurseStr = if ($Recurse) { ' -Recurse' } else { '' }
    Write-Host "  [exec] Copy-Item$recurseStr '$Path' -> '$Destination'" -ForegroundColor DarkGray
    $details = @{ Operation = 'Copy-Item'; Path = $Path; Destination = $Destination }
    Invoke-Logged -Description 'Copy-Item' -Details $details -ScriptBlock {
        Copy-Item -Path $Path -Destination $Destination -Recurse:$Recurse -Force:$Force -ErrorAction SilentlyContinue
    }
}


# Helper: Detect partition role (Windows, Boot, or Unknown)
function Get-PartitionRole {
    param(
        [Parameter(Mandatory = $true)][string]$AccessPath,
        [Parameter(Mandatory = $true)]$PartitionInfo
    )
    
    $role = @()
    
    # Check for Windows partition
    $winPath = Join-Path $AccessPath "Windows\System32\ntdll.dll"
    if (Test-Path $winPath) {
        $role += "Windows"
    }
    
    # Check for UEFI/GPT Boot partition.
    # Primary check: BCD file present. Secondary: GPT System partition with EFI folder tree
    # (covers the case where BCD was deleted but the partition structure is intact).
    $efiPathBcd = Join-Path $AccessPath "EFI\Microsoft\Boot\BCD"
    $efiPathDir = Join-Path $AccessPath "EFI\Microsoft\Boot"
    $isEfiPartition = $PartitionInfo.Type -eq 'System'
    if ($isEfiPartition -and (Test-Path $efiPathBcd)) {
        $role += "Boot (UEFI)"
    } elseif ($isEfiPartition -and (Test-Path $efiPathDir)) {
        $role += "Boot (UEFI - BCD missing)"
    } elseif ($isEfiPartition) {
        $role += "Boot (UEFI - EFI folder missing)"
    }
    
    # Check for BIOS/MBR Boot partition
    $biosPathBcd = Join-Path $AccessPath "Boot\BCD"
    $biosPathDir = Join-Path $AccessPath "Boot"
    $isActivePartition = $PartitionInfo.IsActive
    if ($isActivePartition -and (Test-Path $biosPathBcd)) {
        $role += "Boot (BIOS)"
    } elseif ($isActivePartition -and (Test-Path $biosPathDir)) {
        $role += "Boot (BIOS - BCD missing)"
    } elseif ($isActivePartition) {
        $role += "Boot (BIOS - Boot folder missing)"
    }
    
    if ($role.Count -eq 0) {
        return "Unknown"
    }
    return $role -join " + "
}

# Helper: Detect disk generation (BIOS vs UEFI) based on partitions
function Get-DiskGeneration {
    param(
        [Parameter(Mandatory = $true)]$Disk
    )
    
    # Check partition table style
    $partStyle = $Disk.PartitionStyle
    
    if ($partStyle -eq 'GPT') {
        return 2  # UEFI/GPT = Gen2
    }
    elseif ($partStyle -eq 'MBR') {
        return 1  # BIOS/MBR = Gen1
    }
    
    return $null
}

# Helper: Executes a PowerShell command as the SYSTEM account using a temporary scheduled task.
function ExecuteAsSystem($cmd) {
    Write-Host "  [exec] ExecuteAsSystem: $cmd" -ForegroundColor DarkGray
    Try {
        $taskName = "TempSystemTask_$([guid]::NewGuid().ToString())"
        # Encode the command as Base64 and use -EncodedCommand so that special
        # characters in the command string (e.g. quotes, spaces in paths) cannot
        # break out of the argument and inject arbitrary code.
        $encodedCmd = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($cmd))
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -EncodedCommand $encodedCmd"
        $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
        $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1)

        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal | Out-Null

        Start-ScheduledTask -TaskName $taskName

        While ((Get-ScheduledTask -TaskName $taskName).State -eq "Running" -Or (Get-ScheduledTaskInfo -TaskName $taskName).LastRunTime -notmatch (Get-Date).ToString("yyyy")) {
            Start-Sleep -Seconds 1
        }

        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false | Out-Null
    }
    Catch {
        Write-Error "Failed to execute command as SYSTEM: $_"
    }
}

function MountOffHive {
    param(
        [string] $WinPath,
        [string] $Hive
    )

    if ($Hive -eq "SYSTEM") {
        $OffHive = Join-Path $WinPath 'System32\Config\SYSTEM'
    }
    elseif ($Hive -eq "SOFTWARE") {
        $OffHive = Join-Path $WinPath 'System32\Config\SOFTWARE'
    }
    elseif ($Hive -eq "COMPONENTS") {
        $OffHive = Join-Path $WinPath 'System32\Config\COMPONENTS'
    }    
    
    if (-not (Test-Path $OffHive)) { throw "$Hive hive not found: $OffHive" }
    Write-Host "reg load HKLM\BROKEN$($Hive) `"$OffHive`""
    $out = reg.exe load HKLM\BROKEN$($Hive) "$OffHive" 2>&1
    if ($LASTEXITCODE -ne 0 -and $out -notmatch 'error: The specified registry key is already loaded') {
        throw "Failed to load offline $Hive hive: $out"
    }      
}

function UnmountOffHive {
    param(
        [string] $Hive
    )

    # Try to release PowerShell/.NET references
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 200

    Write-Host "reg unload HKLM\BROKEN$($Hive)"
    $null = reg.exe unload HKLM\BROKEN$($Hive) 2>&1
    Start-Sleep -Seconds 2
}

# Helper: Disable service or driver
function Disable-ServiceOrDriver {
    param(
        [Parameter(Mandatory = $true)]$ServiceName
    )

    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        $SystemRoot      = Get-SystemRootPath
        $ServiceFullPath = "$SystemRoot\Services\$ServiceName"

        if (Test-Path $ServiceFullPath) {
            $CurrentValue = (Get-ItemProperty -Path $ServiceFullPath -Name Start).Start
            Write-Host "Current $ServiceName Start -> $CurrentValue`r`nSetting to 4."
            Set-ItemProperty-Logged -Path $ServiceFullPath -Name Start -Value 4 -Type DWord -Force
        }
        else {
            Write-Host "Service path $ServiceFullPath not found."
        }
    }
    catch {
        Write-Error "Disable-ServiceOrDriver failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

# Helper: Enable service or driver
function Enable-ServiceOrDriver {
    param(
        [Parameter(Mandatory = $true)]$ServiceName,
        [Parameter(Mandatory = $true)]$StartValue
    )

    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        $SystemRoot      = Get-SystemRootPath
        $ServiceFullPath = "$SystemRoot\Services\$ServiceName"

        if (Test-Path $ServiceFullPath) {
            $CurrentValue = (Get-ItemProperty -Path $ServiceFullPath -Name Start -ErrorAction SilentlyContinue).Start
            if ($null -ne $CurrentValue -and [int]$CurrentValue -eq [int]$StartValue) {
                Write-Host "  $ServiceName Start = $CurrentValue (already correct, no change needed)." -ForegroundColor DarkGray
            } else {
                Write-Host "  $ServiceName Start: $CurrentValue -> $StartValue" -ForegroundColor Cyan
                Set-ItemProperty-Logged -Path $ServiceFullPath -Name Start -Value $StartValue -Type DWord -Force
            }
        }
        else {
            Write-Host "Service path $ServiceFullPath not found." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "Enable-ServiceOrDriver failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

# Helper: Enable registry periodic backups (regback)
function EnableRegBackup {
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        $SystemRoot = Get-SystemRootPath

        Write-Host "Enabling registry backup feature..." #https://learn.microsoft.com/en-us/troubleshoot/windows-client/installing-updates-features-roles/system-registry-no-backed-up-regback-folder
        Set-ItemProperty-Logged -Path "$SystemRoot\Control\Session Manager\Configuration Manager" -Name EnablePeriodicBackup -Value 1 -Type DWord -Force
    }
    catch {
        Write-Error "EnableRegBackup failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

# Helper to create key + set DWORD value (uses Set-ItemProperty-Logged for full audit trail)
function Set-DwordValue {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][int]$Value
    )
    Set-ItemProperty-Logged -Path $Path -Name $Name -Value $Value -Type DWord -Force

    [PSCustomObject]@{
        RegistryPath = $Path
        Name         = $Name
        Value        = $Value
    }
}

################################################################################
#+ Operational functions (moved from previous inline flow)
################################################################################

# Helper: Asks the user to confirm before running an operation that makes
# significant changes (large file deletions, hive overwrites, bulk driver changes).
# Returns $true if the user types 'Y', $false to skip with no changes made.
function Confirm-CriticalOperation {
    param(
        [Parameter(Mandatory)][string]$Operation,
        [Parameter(Mandatory)][string]$Details
    )

    Write-Host ""
    Write-Host "  $Operation" -ForegroundColor Cyan
    Write-Host "  What this will do:" -ForegroundColor DarkGray
    $Details -split "`n" | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray }
    Write-Host "  Note: the live hives / files are backed up first where possible, but" -ForegroundColor DarkGray
    Write-Host "  having a VM snapshot before running this is always a good idea." -ForegroundColor DarkGray
    Write-Host ""

    $answer = Read-Host "  Continue? [Y] Yes  [N] No (default: No)"
    if ($answer -ne 'Y' -and $answer -ne 'y') {
        Write-Host "  Skipped. No changes were made." -ForegroundColor DarkGray
        return $false
    }
    return $true
}

function Get-SystemRootPath {
    # Returns the active ControlSet path in the mounted BROKENSYSTEM hive.
    # Falls back to ControlSet001 when the Select key is absent (e.g. offline WinPE disks).
    $current = (Get-ItemProperty 'HKLM:\BROKENSYSTEM\Select' -ErrorAction SilentlyContinue).Current
    if ($current) { return 'HKLM:\BROKENSYSTEM\ControlSet{0:d3}' -f $current }
    return 'HKLM:\BROKENSYSTEM\ControlSet001'
}

function Ensure-GuestTempDir {
    # Creates the \Temp directory on the offline guest disk if it does not already exist.
    if (-not (Test-Path -Path "$script:WinDriveLetter\Temp")) {
        New-Item-Logged -Path "$script:WinDriveLetter\Temp" -ItemType Directory -Force | Out-Null
    }
}

function FixDiskCorruption {
    param([string]$DriveLetter = '')
    $target = if ($DriveLetter) { $DriveLetter.TrimEnd('\') } else { $script:WinDriveLetter.TrimEnd('\') }
    Write-Host "Running chkdsk on $target ..." -ForegroundColor Yellow
    try {
        & chkdsk $target /F /X
    }
    catch {
        Write-Error "FixDiskCorruption failed: $_"
    }
}

function RunDismHealth {
    Write-Host "Running DISM /ScanHealth and /RestoreHealth" -ForegroundColor Yellow
    try {
        if (-not (Test-Path "C:\temp")) { mkdir C:\temp | Out-Null }
        & dism /Image:$script:WinDriveLetter /Cleanup-Image /ScanHealth /ScratchDir:C:\Temp
        & dism /Image:$script:WinDriveLetter /Cleanup-Image /RestoreHealth /Source:C:\Windows\WinSxS /LimitAccess /ScratchDir:C:\Temp
    }
    catch {
        Write-Error "RunDismHealth failed: $_"
        throw
    }
}

function RunSFC {
    Write-Host "Executing SFC /scannow"
    try {
        $sfcCmd = "sfc /SCANNOW /OFFBOOTDIR=$script:WinDriveLetter /OFFWINDIR=$($script:WinDriveLetter)windows"
        Write-Host "  [exec] $sfcCmd" -ForegroundColor DarkGray
        & sfc /SCANNOW /OFFBOOTDIR=$script:WinDriveLetter /OFFWINDIR=$($script:WinDriveLetter)windows
    }
    catch {
        Write-Error "RunSFC failed: $_"
    }
}

function SetLKGC {
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"

    try {
        $CurrentBoot = (Get-ItemProperty -Path "HKLM:\BROKENSYSTEM\Select" -Name Current).Current
        $LastKnownGood = (Get-ItemProperty -Path "HKLM:\BROKENSYSTEM\Select" -Name LastKnownGood).LastKnownGood
        Write-Host "Current HKLM: $CurrentBoot`r`nLast Known Good: $LastKnownGood" -ForegroundColor Green
        Write-Host "`r`nSetting next boot to LKGD: $LastKnownGood" -ForegroundColor Yellow
        Set-ItemProperty-Logged -Path "HKLM:\BROKENSYSTEM\Select" -Name Current -Value $LastKnownGood -Type DWord -Force
    }
    catch {
        Write-Error "SetLKGC failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

function RevertLKGC {
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"

    try {
        $CurrentBoot = (Get-ItemProperty -Path "HKLM:\BROKENSYSTEM\Select" -Name Current).Current
        Write-Host "Current HKLM: $CurrentBoot" -ForegroundColor Green

        $NextSetting = Get-ChildItem 'HKLM:\BROKENSYSTEM' |
        Where-Object { $_.PSChildName -like 'ControlSet*' } |
        ForEach-Object {
            [int]($_.PSChildName -replace 'ControlSet00', '')
        } |
        Where-Object { $_ -ne $CurrentBoot } |
        Select-Object -First 1

        Write-Host "`r`nSetting boot registry to: $NextSetting" -ForegroundColor Yellow
        Set-ItemProperty-Logged -Path "HKLM:\BROKENSYSTEM\Select" -Name Current -Value $NextSetting -Type DWord -Force
    }
    catch {
        Write-Error "RevertLKGC failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

function SetBootLog {
    $storePath = Get-BcdStorePath -Generation $script:VMGen -BootDrive $script:BootDriveLetter
    $identifier = Get-BcdBootLoaderId -StorePath $storePath
    
    if (-not $identifier) { return }
    
    Write-Host "Enabling boot logging..." -ForegroundColor Green

    $NtBtLog = Join-Path $script:WinDriveLetter "Windows\ntbtlog.txt"
    if (Test-Path $NtBtLog) {
        Write-Host "Deleting existing ntbtlog.txt..." -ForegroundColor Yellow
        Remove-Item-Logged -Path $NtBtLog -Force
    }

    Invoke-BcdEdit -StorePath $storePath -Command "/set $identifier bootlog yes"
}

function SetSafeMode {
    $storePath = Get-BcdStorePath -Generation $script:VMGen -BootDrive $script:BootDriveLetter
    $identifier = Get-BcdBootLoaderId -StorePath $storePath
    
    if (-not $identifier) { return }
    
    Write-Host "Setting Safe Mode minimal flag..." -ForegroundColor Green
    Invoke-BcdEdit -StorePath $storePath -Command "/set $identifier safeboot minimal"
}

function RemoveSafeMode {
    $storePath = Get-BcdStorePath -Generation $script:VMGen -BootDrive $script:BootDriveLetter
    $identifier = Get-BcdBootLoaderId -StorePath $storePath
    
    if (-not $identifier) { return }
    
    Write-Host "Removing Safe Mode flag..."
    Invoke-BcdEdit -StorePath $storePath -Command "/deletevalue $identifier safeboot"
}

function ConfigureFullMemDump {
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        $SystemRoot = Get-SystemRootPath

        Write-Host "Configuring full memory dump settings..." -ForegroundColor Yellow
        Set-ItemProperty-Logged -Path "$SystemRoot\Control\CrashControl" -Name CrashDumpEnabled -Value 1 -Type DWord -Force
        Set-ItemProperty-Logged -Path "$SystemRoot\Control\CrashControl" -Name DumpFile -Value "C:\Windows\MEMORY.dmp" -Type ExpandString -Force
        Set-ItemProperty-Logged -Path "$SystemRoot\Control\CrashControl" -Name DedicatedDumpFile -Value "C:\DD.sys" -Type String -Force
        Set-ItemProperty-Logged -Path "$SystemRoot\Control\CrashControl" -Name Overwrite -Value 1 -Type DWord -Force
        Set-ItemProperty-Logged -Path "$SystemRoot\Control\CrashControl" -Name NMICrashDump -Value 1 -Type DWord -Force
        Set-ItemProperty-Logged -Path "$SystemRoot\Control\CrashControl" -Name AutoReboot -Value 1 -Type DWord -Force

        Write-Host "Check for dump as 'C:\Windows\MEMORY.dmp'." -ForegroundColor Green

        $CurrentPagefile = (Get-ItemProperty -Path "$SystemRoot\Control\Session Manager\Memory Management" -Name PagingFiles).PagingFiles
        if ($CurrentPagefile -notmatch "C:") {
            Write-Host "Configuring pagefile to C: drive..." -ForegroundColor Yellow
            Set-ItemProperty-Logged -Path "$SystemRoot\Control\Session Manager\Memory Management" -Name PagingFiles -Value "C:\pagefile.sys 0 0" -Type MultiString -Force
        }
    }
    catch {
        Write-Error "ConfigureFullMemDump failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

function RebuildBCD {
    try {
        $storePath = Get-BcdStorePath -Generation $script:VMGen -BootDrive $script:BootDriveLetter

        # Backup existing BCD if present
        if (Test-Path -LiteralPath $storePath) {
            $bakPath = New-UniqueBackupPath -BasePath $storePath -BakSuffix ".bak"
            Write-Host "Renaming BCD: $storePath -> $bakPath"
            Move-Item-Logged -LiteralPath $storePath -Destination $bakPath -Force
        }
        else {
            Write-Warning "BCD not found at $storePath; will still run bcdboot to create a fresh store."
        }

        # Prepare drive letters (strip trailing backslashes for bcdboot syntax)
        $WinDrive = $script:WinDriveLetter.TrimEnd('\')
        $SysDrive = $script:BootDriveLetter.TrimEnd('\')

        # Rebuild BCD from scratch
        $format = if ($script:VMGen -eq 1) { "BIOS" } else { "UEFI" }
        $rebuildCmd = "bcdboot $WinDrive\Windows /s $SysDrive /v /f $format"
        Write-Host "Rebuilding BCD (Gen$script:VMGen): $rebuildCmd"
        & cmd.exe /c $rebuildCmd

        # Verify rebuild
        $verifyCmd = "bcdedit /store `"$storePath`" /enum"
        Write-Host "Verifying: $verifyCmd"
        & cmd.exe /c $verifyCmd | Out-Host

        # Get boot loader identifier
        $identifier = Get-BcdBootLoaderId -StorePath $storePath
        if (-not $identifier) { return }

        # Apply additional boot configuration flags (same for both Gen1 and Gen2)
        $extraCmds = @(
            "/set $identifier integrityservices enable",
            "/set $identifier recoveryenabled Off",
            "/set $identifier bootstatuspolicy IgnoreAllFailures",
            "/set {bootmgr} displaybootmenu yes",
            "/set {bootmgr} timeout 5",
            "/set {bootmgr} bootems yes",
            "/ems $identifier on",
            "/emssettings EMSPORT:1 EMSBAUDRATE:115200"
        )

        foreach ($cmd in $extraCmds) {
            Write-Host "Applying extra setting: $cmd" -ForegroundColor Yellow
            Invoke-BcdEdit -StorePath $storePath -Command $cmd
        }
    }
    catch {
        Write-Error "RebuildBCD failed: $_"
        throw
    }
}

function AddNewUser {
    Param (
        [Parameter(Mandatory = $True)][string]$WinDrive,
        [Parameter(Mandatory = $True)][pscredential]$Credential
    )

    If ($WinDrive) {
        If (-Not $WinDrive.Contains(":")) {
            $WinDrive += ":"
        }

        If (-Not $WinDrive.EndsWith("\")) {
            $WinDrive += "\"
        }
    }

    If (-Not $Credential) {
        $Credential = Get-Credential
    }

    $Username = $Credential.GetNetworkCredential().UserName
    $Password = $Credential.GetNetworkCredential().Password

    $GptIni = @"
[General]
gPCFunctionalityVersion=2
gPCMachineExtensionNames=[{42B5FAAE-6536-11D2-AE5A-0000F87571E3}{40B6664F-4972-11D1-A7CA-0000F87571E3}]
Version=1    
"@

    $FixAzureVM = @"
net user Username "Password" /add /Y
net localgroup administrators Username /add
net localgroup "Remote Desktop Users" Username /add
del /F C:\Windows\System32\GroupPolicy\Machine\Scripts\scripts.ini
del /F C:\Windows\System32\GroupPolicy\gpt.ini
del /F C:\Windows\System32\GroupPolicy\Machine\Scripts\Startup\FixAzureVM.cmd
"@

    $ScriptsIni = @"
[Startup]
0CmdLine=FixAzureVM.cmd
0Parameters=
"@

    $System32Path = Join-Path $WinDrive "Windows\System32"

    If (-Not (Test-Path -Path "$System32Path\GroupPolicy")) {
        New-Item-Logged -Path "$System32Path\GroupPolicy" -ItemType Directory -Force
    }

    If (-Not (Test-Path -Path "$System32Path\GroupPolicy\Scripts")) {
        New-Item-Logged -Path "$System32Path\GroupPolicy\Scripts" -ItemType Directory -Force
    }

    If (-Not (Test-Path -Path "$System32Path\GroupPolicy\Scripts\Startup")) {
        New-Item-Logged -Path "$System32Path\GroupPolicy\Scripts\Startup" -ItemType Directory -Force
    }

    #Create gpt.ini
    New-Item-Logged -Path "$System32Path\GroupPolicy" -Name "gpt.ini" -ItemType File -Value $GptIni.ToString().Trim() -Force
    #Create scripts.ini
    New-Item-Logged -Path "$System32Path\GroupPolicy\Machine\Scripts" -Name "scripts.ini" -ItemType File -Value $ScriptsIni.ToString().Trim() -Force
    #Create FixAzureVM.cmd
    New-Item-Logged -Path "$System32Path\GroupPolicy\Machine\Scripts\Startup" -Name "FixAzureVM.cmd" -ItemType File -Value $FixAzureVM.ToString().Trim().Replace("Username", $Username).Replace("Password", $Password) -Force -RedactValue

    Write-Host "If the VM is domain joined, local GPO may not apply, try -AddTempUser2 if this doesn't work."
}

function AddNewUser2 {
    param (
        [Parameter(Mandatory = $True)][pscredential]$Credential
    )

    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        Write-Host "Configuring setup and adding script..." -ForegroundColor Green
        Set-ItemProperty-Logged -Path "HKLM:\BROKENSYSTEM\Setup" -Name SetupType -Value 2 -Type DWord -Force
        Set-ItemProperty-Logged -Path "HKLM:\BROKENSYSTEM\Setup" -Name CmdLine -Value "cmd.exe /c c:\temp\adduser.cmd" -Type String -Force

        If (-Not $Credential) {
            $Credential = Get-Credential
        }

        $Username = $Credential.GetNetworkCredential().UserName
        $Password = $Credential.GetNetworkCredential().Password

        $NewUserScript = @"
@echo off
net user Username "Password" /add /Y
net localgroup administrators Username /add
net localgroup "Remote Desktop Users" Username /add
reg add "HKLM\SYSTEM\Setup" /v SetupType /t REG_DWORD /d 0 /f
reg add "HKLM\SYSTEM\Setup" /v CmdLine /t REG_SZ /d "" /f
del /F C:\Temp\adduser.cmd > NUL
"@

        Ensure-GuestTempDir
        New-Item-Logged -Path "$script:WinDriveLetter\Temp" -Name "adduser.cmd" -ItemType File -Value $NewUserScript.ToString().Trim().Replace("Username", $Username).Replace("Password", $Password) -Force -RedactValue

        Write-Host "A script to add the user $Username will run at next boot. If the VM is domain joined, this should work even if local GPO is blocked." -ForegroundColor Green
    }
    catch {
        Write-Error "AddNewUser2 failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}


function GetRdpAuthPolicySnapshot {
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"

    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SOFTWARE"
    try {
        $SystemRoot   = Get-SystemRootPath
        $SoftwareRoot = "HKLM:\BROKENSOFTWARE"

        $checks = @(
            # --- RDP listener ---
            @{ Path = "$SystemRoot\Control\Terminal Server\WinStations\RDP-Tcp"; Name = "SecurityLayer" },
            @{ Path = "$SystemRoot\Control\Terminal Server\WinStations\RDP-Tcp"; Name = "UserAuthentication" },
            @{ Path = "$SystemRoot\Control\Terminal Server\WinStations\RDP-Tcp"; Name = "SSLCertificateSHA1Hash" },

            # --- NTLM restrictions ---
            @{ Path = "$SystemRoot\Control\Lsa\MSV1_0"; Name = "RestrictReceivingNTLMTraffic" },
            @{ Path = "$SystemRoot\Control\Lsa\MSV1_0"; Name = "RestrictSendingNTLMTraffic" },
            @{ Path = "$SystemRoot\Control\Lsa"; Name = "LmCompatibilityLevel" },

            # --- CredSSP / Encryption Oracle ---
            @{ Path = "$SoftwareRoot\Microsoft\Windows\CurrentVersion\Policies\System\CredSSP\Parameters"; Name = "AllowEncryptionOracle" }
        )

        foreach ($c in $checks) {
            $val = $null
            $status = "Missing"

            try {
                if (Test-Path $c.Path) {
                    $p = Get-ItemProperty -Path $c.Path -ErrorAction Stop
                    if ($null -ne $p.($c.Name)) {
                        $val = $p.($c.Name)
                        $status = "Present"
                    }
                }

                # Special handling: SSLCertificateSHA1Hash is REG_BINARY -> show as hex string
                if ($status -eq "Present" -and $c.Name -eq "SSLCertificateSHA1Hash" -and $val -is [byte[]]) {
                    $val = ([BitConverter]::ToString($val) -replace '-', '')
                }
            }
            catch {
                $status = "Error: $($_.Exception.Message)"
            }

            [PSCustomObject]@{
                RegistryPath = $c.Path
                Name         = $c.Name
                Value        = $val
                Status       = $status
            }
        }
    }
    catch {
        Write-Error "GetRdpAuthPolicySnapshot failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
        UnmountOffHive -Hive "SOFTWARE"
    }
}

function SetRdpAuthPolicyOptimal {
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"

    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SOFTWARE"

    try {
        $SystemRoot   = Get-SystemRootPath
        $SoftwareRoot = "HKLM:\BROKENSOFTWARE"

        # --- Define the changes to apply: Path, Name, new Value ---
        $rdpTcp  = "$SystemRoot\Control\Terminal Server\WinStations\RDP-Tcp"
        $msv10   = "$SystemRoot\Control\Lsa\MSV1_0"
        $lsa     = "$SystemRoot\Control\Lsa"
        $credssp = "$SoftwareRoot\Microsoft\Windows\CurrentVersion\Policies\System\CredSSP\Parameters"

        $changes = @(
            @{ Path = $rdpTcp;  Name = "SecurityLayer";                 Value = 2 },
            @{ Path = $rdpTcp;  Name = "UserAuthentication";            Value = 0 },
            @{ Path = $msv10;   Name = "RestrictReceivingNTLMTraffic";  Value = 0 },
            @{ Path = $msv10;   Name = "RestrictSendingNTLMTraffic";    Value = 0 },
            @{ Path = $lsa;     Name = "LmCompatibilityLevel";          Value = 3 },
            @{ Path = $credssp; Name = "AllowEncryptionOracle";         Value = 2 }
        )

        # --- Snapshot current values before making any changes ---
        $snapshots = foreach ($c in $changes) {
            $before = $null
            $existed = $false
            if (Test-Path $c.Path) {
                $prop = Get-ItemProperty -Path $c.Path -Name $c.Name -ErrorAction SilentlyContinue
                if ($null -ne $prop.($c.Name)) {
                    $before  = $prop.($c.Name)
                    $existed = $true
                }
            }
            [PSCustomObject]@{
                Path    = $c.Path
                Name    = $c.Name
                NewValue = $c.Value
                Before  = $before
                Existed = $existed
            }
        }

        # --- Apply changes ---
        $results = @()
        foreach ($c in $changes) {
            $results += Set-DwordValue -Path $c.Path -Name $c.Name -Value $c.Value
        }

        Write-Host "`nApplied the following RDP auth policy changes:" -ForegroundColor Green
        $results | Format-Table -AutoSize

        Write-Host "RDP will use TLS security layer but allow fallback to non-NLA auth, NTLM auth will be allowed with NTLMv2-only responses, and CredSSP will allow vulnerable encryption oracle for compatibility." -ForegroundColor Yellow
        Write-Host "These settings are for recovery purposes only. After the VM is accessible again, restore the original values using the commands below." -ForegroundColor Yellow

        # --- Print restore commands based on the captured snapshots ---
        Write-Host "`n--- REVERT COMMANDS (run these on the VM after recovery to restore original settings) ---" -ForegroundColor Cyan
        Write-Host "# Mount the offline hives first if reverting via the repair host, or run directly on the live VM." -ForegroundColor DarkCyan
        foreach ($s in $snapshots) {
            # Convert the BROKEN hive path back to the live registry path for the restore commands
            $livePath = $s.Path -replace '^HKLM:\\BROKENSYSTEM\\', 'HKLM:\SYSTEM\' `
                                -replace '^HKLM:\\BROKENSOFTWARE\\', 'HKLM:\SOFTWARE\'
            if ($s.Existed) {
                Write-Host "Set-ItemProperty -Path '$livePath' -Name '$($s.Name)' -Value $($s.Before) -Type DWord -Force" -ForegroundColor White
            }
            else {
                Write-Host "Remove-ItemProperty -Path '$livePath' -Name '$($s.Name)' -ErrorAction SilentlyContinue  # value did not exist before" -ForegroundColor DarkGray
            }
        }
        Write-Host "-------------------------------------------------------------------------------------`n" -ForegroundColor Cyan
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
        UnmountOffHive -Hive "SOFTWARE"
    }
}

function ResetRDPSettings {
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"

    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SOFTWARE"

    $SystemRoot   = Get-SystemRootPath
    $SoftwareRoot = "HKLM:\BROKENSOFTWARE"

    $TSKeyPath    = "$SystemRoot\Control\Terminal Server"
    $WinStationsPath = "$SystemRoot\Control\Terminal Server\Winstations"
    $RdpTcpPath = "$SystemRoot\Control\Terminal Server\Winstations\RDP-Tcp"
    $TSPolicyPath = "$SoftwareRoot\Policies\Microsoft\Windows NT\Terminal Services"

    try {    
        Write-Host "Setting RDP to default configuration..." -ForegroundColor Yellow

        # ── Enable RDP at the canonical registry key ──────────────────────────
        # fDenyTSConnections=0 must be set on the Terminal Server key itself,
        # not only on the policy path, otherwise RDP remains blocked.
        Set-ItemProperty-Logged -Path $TSKeyPath -Name fDenyTSConnections -Value 0 -Type Dword -Force

        # ── RDP-dependent services ────────────────────────────────────────────
        # TermService (Remote Desktop Services) - must not be disabled.
        # SessionEnv  (Remote Desktop Config)   - required for TermService to start.
        # UmRdpService (Remote Desktop UserMode Port Redirector) - required for redirectors.
        # All three default to Manual (3); set to Auto (2) so they survive reboots reliably.
        $rdpServices = @(
            @{ Name = 'TermService';   DefaultStart = 2 }
            @{ Name = 'SessionEnv';    DefaultStart = 2 }
            @{ Name = 'UmRdpService';  DefaultStart = 2 }
        )
        foreach ($svc in $rdpServices) {
            $svcPath = "$SystemRoot\Services\$($svc.Name)"
            if (Test-Path $svcPath) {
                $current = (Get-ItemProperty $svcPath -ErrorAction SilentlyContinue).Start
                if ($current -eq 4 -or $null -eq $current) {
                    Write-Host "  $($svc.Name): Start was $(if ($null -eq $current) {'(not set)'} else {'Disabled (4)'}) -> setting to Auto (2)" -ForegroundColor Yellow
                    Set-ItemProperty-Logged -Path $svcPath -Name Start -Value $svc.DefaultStart -Type DWord -Force
                } else {
                    Write-Host "  $($svc.Name): Start=$current (no change needed)" -ForegroundColor DarkGray
                }
            } else {
                Write-Host "  $($svc.Name): service key not found - skipping" -ForegroundColor DarkGray
            }
        }

        Set-ItemProperty-Logged -Path $WinStationsPath -Name SelfSignedCertStore -Value 'Remote Desktop' -Type String -Force

        Set-ItemProperty-Logged -Path $RdpTcpPath -Name PortNumber -Value 3389 -Type Dword -Force
        Set-ItemProperty-Logged -Path $RdpTcpPath -Name fInheritReconnectSame -Value 1 -Type Dword -Force
        Set-ItemProperty-Logged -Path $RdpTcpPath -Name fReconnectSame -Value 1 -Type Dword -Force
        Set-ItemProperty-Logged -Path $RdpTcpPath -Name fInheritMaxSessionTime -Value 1 -Type Dword -Force
        Set-ItemProperty-Logged -Path $RdpTcpPath -Name fInheritMaxDisconnectionTime -Value 1 -Type Dword -Force
        Set-ItemProperty-Logged -Path $RdpTcpPath -Name MaxDisconnectionTime -Value 0 -Type Dword -Force
        Set-ItemProperty-Logged -Path $RdpTcpPath -Name MaxConnectionTime -Value 0 -Type Dword -Force
        Set-ItemProperty-Logged -Path $RdpTcpPath -Name fInheritMaxIdleTime -Value 1 -Type Dword -Force
        Set-ItemProperty-Logged -Path $RdpTcpPath -Name MaxIdleTime -Value 0 -Type Dword -Force
        Set-ItemProperty-Logged -Path $RdpTcpPath -Name MaxInstanceCount -Value 4294967295 -Type Dword -Force
        Set-ItemProperty-Logged -Path $RdpTcpPath -Name LanAdapter -Value 0 -Type Dword -Force
        Set-ItemProperty-Logged -Path $RdpTcpPath -Name TSServerDrainMode -Value 0 -Type Dword -Force
        Set-ItemProperty-Logged -Path $RdpTcpPath -Name fQueryUserConfigFromLocalMachine -Value 1 -Type Dword -Force

        Set-ItemProperty-Logged -Path $TSPolicyPath -Name KeepAliveEnable -Value 1 -Type Dword -Force
        Set-ItemProperty-Logged -Path $TSPolicyPath -Name KeepAliveInterval -Value 1 -Type Dword -Force
        Set-ItemProperty-Logged -Path $TSPolicyPath -Name KeepAliveTimeout -Value 1 -Type Dword -Force
        Set-ItemProperty-Logged -Path $TSPolicyPath -Name fDenyTSConnections -Value 0 -Type Dword -Force
        Set-ItemProperty-Logged -Path $TSPolicyPath -Name fDisableAutoReconnect -Value 0 -Type Dword -Force

        $SSLPolicyPath = "HKLM:\BROKENSOFTWARE\Policies\Microsoft\Cryptography\Configuration\SSL\00010002"
        Write-Host "Clearing SSL 00010002 'Functions'..." -ForegroundColor Yellow
        if (Test-Path -Path $SSLPolicyPath) {
            Remove-ItemProperty-Logged -Path $SSLPolicyPath -Name Functions -Force -ErrorAction SilentlyContinue
        }
        else {
            New-Item-Logged -Path $SSLPolicyPath -Force
        }

        $tls12Paths = @(
            "$SystemRoot\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client",
            "$SystemRoot\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server"
        )

        foreach ($tls12Path in $tls12Paths) {
            if (-not (Test-Path $tls12Path)) {
                New-Item-Logged -Path $tls12Path -Force
            }

            Set-ItemProperty-Logged -Path $tls12Path -Name Enabled -Value 1 -Type DWord -Force | Out-Null
            Set-ItemProperty-Logged -Path $tls12Path -Name DisabledByDefault -Value 0 -Type DWord -Force | Out-Null
        }
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
        UnmountOffHive -Hive "SOFTWARE"
    }    
}

function ResetRDPPrivKeyPermissions {
    Write-Host "Resetting Windows RSA MachineKeys permissions..." -ForegroundColor Green
    $machineKeysPath = "$($script:WinDriveLetter)ProgramData\Microsoft\Crypto\RSA\MachineKeys"
    Write-Host "  [exec] takeown /f '$machineKeysPath' /a /r" -ForegroundColor DarkGray
    & takeown /f $machineKeysPath /a /r | Out-Null
    Write-Host "  [exec] icacls '$machineKeysPath' /t /c /grant 'NT AUTHORITY\SYSTEM:(F)'" -ForegroundColor DarkGray
    & icacls $machineKeysPath /t /c /grant "NT AUTHORITY\SYSTEM:(F)" | Out-Null
    Write-Host "  [exec] icacls '$machineKeysPath' /t /c /grant 'NT AUTHORITY\NETWORK SERVICE:(R)'" -ForegroundColor DarkGray
    & icacls $machineKeysPath /t /c /grant "NT AUTHORITY\NETWORK SERVICE:(R)" | Out-Null
    Write-Host "  [exec] icacls '$machineKeysPath' /t /c /grant 'BUILTIN\Administrators:(F)'" -ForegroundColor DarkGray
    & icacls $machineKeysPath /t /c /grant "BUILTIN\Administrators:(F)" | Out-Null

    $PrivKeys = Get-ChildItem -Path "$($script:WinDriveLetter)ProgramData\Microsoft\Crypto\RSA\MachineKeys\f686aace6942fb7f7ceb231212eef4a4*"

    Write-Host "Resetting RDP certificate private key permissions..." -ForegroundColor Green

    ForEach ($PrivKey in $PrivKeys) {
        $pkPath = $PrivKey.FullName
        Write-Host "  [exec] takeown /f '$pkPath'" -ForegroundColor DarkGray
        & takeown.exe /f $pkPath | Out-Null
        Write-Host "  [exec] icacls '$pkPath' /c /grant 'NT AUTHORITY\SYSTEM:(F)'" -ForegroundColor DarkGray
        & icacls.exe $pkPath /c /grant "NT AUTHORITY\SYSTEM:(F)" | Out-Null
        Write-Host "  [exec] icacls '$pkPath' /c /grant 'NT AUTHORITY\NETWORK SERVICE:(R)'" -ForegroundColor DarkGray
        & icacls.exe $pkPath /c /grant "NT AUTHORITY\NETWORK SERVICE:(R)" | Out-Null
        Write-Host "  [exec] icacls '$pkPath' /c /grant 'NT Service\SessionEnv:(F)'" -ForegroundColor DarkGray
        & icacls.exe $pkPath /c /grant "NT Service\SessionEnv:(F)" | Out-Null
        Write-Host "  [exec] icacls '$pkPath' (display current ACL)" -ForegroundColor DarkGray
        & icacls.exe $pkPath

        ExecuteAsSystem "icacls.exe `"$pkPath`" /setowner 'NT AUTHORITY\SYSTEM'"
    }

    Write-Host "Setting certificate services to default: CertPropSvc, KeyIso, CryptSvc, SessionEnv" -ForegroundColor Yellow
    Enable-ServiceOrDriver -ServiceName CertPropSvc -StartValue 3
    Enable-ServiceOrDriver -ServiceName KeyIso -StartValue 3
    Enable-ServiceOrDriver -ServiceName CryptSvc -StartValue 2
    Enable-ServiceOrDriver -ServiceName SessionEnv -StartValue 3
}

function RecreateRDPCertificate {
    $OfflineWindowsPath = Join-Path $WinDriveLetter "Windows"

    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        $SystemRoot = Get-SystemRootPath

        Write-Host "Configuring setup and adding script..." -ForegroundColor Green
        Set-ItemProperty-Logged -Path "HKLM:\BROKENSYSTEM\Setup" -Name SetupType -Value 2 -Type DWord -Force
        Set-ItemProperty-Logged -Path "HKLM:\BROKENSYSTEM\Setup" -Name CmdLine -Value "cmd.exe /c c:\temp\newrdscert.cmd" -Type String -Force

        $newrdscert = @"
@echo off
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File C:\temp\rdscert.ps1
reg add "HKLM\SYSTEM\Setup" /v SetupType /t REG_DWORD /d 0 /f
reg add "HKLM\SYSTEM\Setup" /v CmdLine /t REG_SZ /d "" /f
del /F C:\temp\rdscert.ps1 > NUL
del /F C:\temp\newrdscert.cmd > NUL
"@

        $RecreateRDPCert = @'
$Passwd =  "REPLACE_WITH_PASSWORD"
$CertPasswd = ConvertTo-SecureString -String $Passwd -Force -AsPlainText
 
#Import certificate into 'Remote Desktop' and 'My' stores
Import-PfxCertificate -FilePath C:\temp\rds.pfx -CertStoreLocation 'Cert:\LocalMachine\Remote Desktop\' -Password $CertPasswd
Import-PfxCertificate -FilePath C:\temp\rds.pfx -CertStoreLocation 'Cert:\LocalMachine\My\' -Password $CertPasswd
 
#Set RDP protocol to use that just imported certificate
wmic /namespace:\\root\cimv2\TerminalServices PATH Win32_TSGeneralSetting Set SSLCertificateSHA1Hash="REPLACE_WITH_THUMBPRINT"

$bytes = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp").SSLCertificateSHA1Hash

$expected = "REPLACE_WITH_THUMBPRINT"
$actual = ([BitConverter]::ToString($bytes) -replace '-', '')
$certbinary = ([byte[]]([regex]::Matches($expected,"..") | % { [Convert]::ToByte($_.Value,16) }))

$expected = $expected.ToUpper()
$actual   = $actual.ToUpper()

    if ($actual -ne $expected) {
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp" -Name SSLCertificateSHA1Hash -Type Binary -Force -Value $certbinary
    }
 
#Restart RDP service (try to restart twice)
Restart-Service TermService -Force
Restart-Service TermService -Force

#Enable RDP and SMB rules in Windows Firewall
Enable-NetFirewallRule FPS-SMB-In-TCP
Enable-NetFirewallRule RemoteDesktop-UserMode-In-TCP

#Clean up
Remove-Item -Path C:\temp\rds.pfx -Force
'@

        Ensure-GuestTempDir
        $guestName = (Get-ItemProperty "$SystemRoot\Control\ComputerName\ComputerName").ComputerName

        Write-Host "Generating new certificate to guest $guestname locally..." -ForegroundColor Green
        $cert = New-SelfSignedCertificate -Type Custom -KeySpec Signature -Subject "CN=$($guestName)" -KeyExportPolicy Exportable -HashAlgorithm sha256 -KeyLength 2048 `
            -CertStoreLocation "Cert:\LocalMachine\My" -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.1") -NotAfter (Get-Date).AddYears(5)

        # Generate a cryptographically random one-time PFX transport password (never hardcoded)
        $rngBytes = [System.Security.Cryptography.RandomNumberGenerator]::GetBytes(24)
        $PfxPassword = [System.Convert]::ToBase64String($rngBytes) -replace '[^A-Za-z0-9]', 'x'
        $certpwd = ConvertTo-SecureString $PfxPassword -AsPlainText -Force

        $hostPfx = Join-Path $WinDriveLetter "Temp\rds.pfx"
        Write-Host "Exporting new cert with thumbprint $($cert.Thumbprint) to $hostPfx..." -ForegroundColor Green
        Export-PfxCertificate -Cert $cert -FilePath $hostPfx -Password $certpwd | Out-Null
        Remove-Item-Logged -Path "Cert:\LocalMachine\My\$($cert.Thumbprint)" -Force

        New-Item-Logged -Path "$WinDriveLetter\Temp" -Name "newrdscert.cmd" -ItemType File -Value $newrdscert.ToString().Trim() -Force
        New-Item-Logged -Path "$WinDriveLetter\Temp" -Name "rdscert.ps1" -ItemType File -Value $RecreateRDPCert.ToString().Trim().Replace("REPLACE_WITH_PASSWORD", $PfxPassword).Replace("REPLACE_WITH_THUMBPRINT", $cert.Thumbprint) -Force -RedactValue

        Write-Host "A script to recreate the RDP certificate will run at next boot. If the RDP service fails to restart after that, you may need to restart the VM." -ForegroundColor Green
    }
    catch {
        Write-Error "RecreateRDPCertificate failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

function SetNLADisabled {
    $OfflineWindowsPath = Join-Path $WinDriveLetter "Windows"

    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SOFTWARE"
    try {
        $SystemRoot   = Get-SystemRootPath
        $SoftwareRoot = "HKLM:\BROKENSOFTWARE"

        $RdpTcpPath   = "$SystemRoot\Control\Terminal Server\Winstations\RDP-Tcp"
        $TSPolicyPath = "$SoftwareRoot\Policies\Microsoft\Windows NT\Terminal Services"

        Write-Host "Disabling NLA..." -ForegroundColor Green

        Set-ItemProperty-Logged -Path $RdpTcpPath -Name UserAuthentication -Value 0 -Type Dword -Force
        Set-ItemProperty-Logged -Path $RdpTcpPath -Name SecurityLayer -Value 0 -Type Dword -Force
        Set-ItemProperty-Logged -Path $RdpTcpPath -Name fAllowSecProtocolNegotiation -Value 0 -Type Dword -Force
        Set-ItemProperty-Logged -Path $RdpTcpPath -Name MinEncryptionLevel -Value 1 -Type Dword -Force
 
        Set-ItemProperty-Logged -Path $TSPolicyPath -Name UserAuthentication -Value 0 -Type Dword -Force
        Set-ItemProperty-Logged -Path $TSPolicyPath -Name SecurityLayer -Value 0 -Type Dword -Force
        Set-ItemProperty-Logged -Path $TSPolicyPath -Name fAllowSecProtocolNegotiation -Value 0 -Type Dword -Force
        Set-ItemProperty-Logged -Path $TSPolicyPath -Name MinEncryptionLevel -Value 1 -Type Dword -Force
    }
    catch {
        Write-Error "SetNLADisabled failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
        UnmountOffHive -Hive "SOFTWARE"
    }
}

function SetNLAEnabled {
    $OfflineWindowsPath = Join-Path $WinDriveLetter "Windows"

    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SOFTWARE"
    try {
        $SystemRoot   = Get-SystemRootPath
        $SoftwareRoot = "HKLM:\BROKENSOFTWARE"

        $RdpTcpPath   = "$SystemRoot\Control\Terminal Server\Winstations\RDP-Tcp"
        $TSPolicyPath = "$SoftwareRoot\Policies\Microsoft\Windows NT\Terminal Services"

        Write-Host "Enabling NLA..." -ForegroundColor Green

        Set-ItemProperty-Logged -Path $RdpTcpPath -Name UserAuthentication -Value 1 -Type Dword -Force
        Set-ItemProperty-Logged -Path $RdpTcpPath -Name SecurityLayer -Value 2 -Type Dword -Force
        Set-ItemProperty-Logged -Path $RdpTcpPath -Name fAllowSecProtocolNegotiation -Value 1 -Type Dword -Force
        Set-ItemProperty-Logged -Path $RdpTcpPath -Name MinEncryptionLevel -Value 2 -Type Dword -Force

        # Clear any policy overrides that would re-disable NLA after reboot.
        # These values may not exist (e.g. on a clean VM), so suppress the error.
        if (Test-Path $TSPolicyPath) {
            foreach ($val in @('UserAuthentication','SecurityLayer','fAllowSecProtocolNegotiation','MinEncryptionLevel')) {
                Remove-ItemProperty-Logged -Path $TSPolicyPath -Name $val -Force -ErrorAction SilentlyContinue
            }
        }
    }
    catch {
        Write-Error "SetNLAEnabled failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
        UnmountOffHive -Hive "SOFTWARE"
    }
}

function ClearPendingUpdates {
    if (-not (Confirm-CriticalOperation -Operation 'Fix Pending Updates (-FixPendingUpdates)' -Details @"
Removes TxR and SMI transaction log files (.blf/.regtrans-ms).
Renames pending.xml in WinSxS.
Clears CBS registry keys (PackagesPending, RebootPending, SessionsPending).
Runs DISM /RevertPendingActions, /StartComponentCleanup, and removes pending packages.
"@)) { return }

    $OfflineWindowsPath = Join-Path $WinDriveLetter "Windows"
    if (-not (Test-Path "C:\temp")) { mkdir C:\temp | Out-Null }

    Write-Host "Saving pending Windows Update packages..." -ForegroundColor Yellow
    $dismText = & dism /Image:$WinDriveLetter /Get-Packages /format:list 2>$null | Out-String -Width 9999
    $blocks = $dismText -split "(\r\n){2,}"

    $packages = foreach ($b in $blocks) {
        $h = @{}
        foreach ($line in ($b -split "`r`n")) {
            if ($line -match '^\s*([^:]+)\s*:\s*(.+)\s*$') {
                $h[$matches[1].Trim()] = $matches[2].Trim()
            }
        }
        if ($h.ContainsKey('Package Identity')) {
            [pscustomobject]$h
        }
    }

    $pendingUpdates = $packages | Where-Object { $_.State -eq 'Pending' } | Select-Object -ExpandProperty 'Package Identity'
    $pendingUpdates | Set-Content -Path C:\Temp\PendingUpdatePackages.txt

    Write-Host "Running DISM /Cleanup-Image /StartComponentCleanup" -ForegroundColor Yellow
    & dism /Image:$WinDriveLetter /Cleanup-Image /StartComponentCleanup

    Write-Host "Running DISM /Cleanup-Image /RevertPendingActions" -ForegroundColor Yellow
    & dism /Image:$WinDriveLetter /Cleanup-Image /RevertPendingActions /ScratchDir:C:\Temp

    $pendingUpdates | ForEach-Object {
        Write-Host "Running package uninstall: $($_)" -ForegroundColor Yellow
        & dism /Image:$WinDriveLetter /Remove-Package /PackageName:$_
    }

    Write-Host "Clearing transactions from TxR folder..." -ForegroundColor Yellow
    $TxRFolder = Join-Path $WinDriveLetter "Windows\system32\config\TxR"
    $BackupFolder = Join-Path $TxRFolder "Backup"
    Get-ChildItem -Path $TxRFolder -Force | ForEach-Object { $_.Attributes = 'Normal' }
    New-Item-Logged -Path $TxRFolder -Name "Backup" -ItemType Directory -Force
    Copy-Item-Logged -Path (Join-Path $TxRFolder '*') -Destination $BackupFolder -Force
    Invoke-Logged -Description 'Remove TxR blf/regtrans files' -Details @{ Path = $TxRFolder } -ScriptBlock {
        Get-ChildItem -Path $TxRFolder -File -Filter *.blf -Force | Remove-Item -Force -ErrorAction SilentlyContinue
        Get-ChildItem -Path $TxRFolder -File -Filter *.regtrans-ms -Force | Remove-Item -Force -ErrorAction SilentlyContinue
    }

    Rename-Item-Logged -Path $TxRFolder -NewName "TxR_OLD"
    New-Item-Logged -Path (Join-Path $WinDriveLetter "Windows\system32\config") -Name "TxR" -ItemType Directory -Force

    Write-Host "Clearing transactions from Config folder..." -ForegroundColor Yellow
    $ConfigFolder = Join-Path $WinDriveLetter "Windows\system32\config"
    $BackupFolder = Join-Path $ConfigFolder "BackupCfg"
    Get-ChildItem -Path $ConfigFolder -Force | ForEach-Object { $_.Attributes = 'Normal' }
    New-Item-Logged -Path $ConfigFolder -Name "BackupCfg" -ItemType Directory -Force
    Copy-Item-Logged -Path (Join-Path $ConfigFolder '*') -Destination $BackupFolder -Force
    Invoke-Logged -Description 'Remove Config blf/regtrans files' -Details @{ Path = $ConfigFolder } -ScriptBlock {
        Get-ChildItem -Path $ConfigFolder -File -Filter *.blf -Force | Remove-Item -Force -ErrorAction SilentlyContinue
        Get-ChildItem -Path $ConfigFolder -File -Filter *.regtrans-ms -Force | Remove-Item -Force -ErrorAction SilentlyContinue
    }

    Write-Host "Clearing transactions from SMI folder..." -ForegroundColor Yellow
    $SMIFolder = Join-Path $WinDriveLetter "Windows\System32\SMI\Store\Machine"
    $BackupFolder = Join-Path $SMIFolder "Backup"
    Get-ChildItem -Path $SMIFolder -Force | ForEach-Object { $_.Attributes = 'Normal' }
    New-Item-Logged -Path $SMIFolder -Name "Backup" -ItemType Directory -Force
    Copy-Item-Logged -Path (Join-Path $SMIFolder '*') -Destination $BackupFolder -Force
    Invoke-Logged -Description 'Remove SMI blf/regtrans files' -Details @{ Path = $SMIFolder } -ScriptBlock {
        Get-ChildItem -Path $SMIFolder -File -Filter *.blf -Force | Remove-Item -Force -ErrorAction SilentlyContinue
        Get-ChildItem -Path $SMIFolder -File -Filter *.regtrans-ms -Force | Remove-Item -Force -ErrorAction SilentlyContinue
    }

    Write-Host "Renaming pending.xmls..." -ForegroundColor Yellow
    Rename-Item-Logged -Path (Join-Path $WinDriveLetter "Windows\WinSxS\pending.xml") -NewName "pending.old"

    Write-Host "Deleting registry keys..." -ForegroundColor Yellow
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SOFTWARE"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "COMPONENTS"
    try {
        $ComponentsReg = "HKLM:\BROKENCOMPONENTS"
        $CbsSoftwareReg = "HKLM:\BROKENSOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing"

        Write-Host "Deleting COMPONENTS registry keys..." -ForegroundColor Yellow
        Remove-ItemProperty-Logged -Path $ComponentsReg -Name "ExecutionState" -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty-Logged -Path $ComponentsReg -Name "PendingXmlIdentifier" -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty-Logged -Path $ComponentsReg -Name "NextQueueEntryIndex" -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty-Logged -Path $ComponentsReg -Name "NextQueueEntryIndexBCDB" -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty-Logged -Path $ComponentsReg -Name "AdvancedInstallersNeedResolving" -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty-Logged -Path $ComponentsReg -Name "StoreDirty" -Force -ErrorAction SilentlyContinue

        Write-Host "Deleting SOFTWARE registry keys..." -ForegroundColor Yellow
        Remove-Item-Logged -Path (Join-Path $CbsSoftwareReg "PackagesPending") -Recurse -Force
        Remove-Item-Logged -Path (Join-Path $CbsSoftwareReg "RebootPending") -Recurse -Force
        Remove-Item-Logged -Path (Join-Path $CbsSoftwareReg "SessionsPending") -Recurse -Force
    }
    catch {
        Write-Error "ClearPendingUpdates (registry cleanup) failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SOFTWARE"
        UnmountOffHive -Hive "COMPONENTS"
    }

    Write-Host "You may want to run SFC and Dism /RestoreHealth with script parameters: -RunSFC -FixHealth." -ForegroundColor Green
}

function SetWinRMHTTPSEnabled {
    $OfflineWindowsPath = Join-Path $WinDriveLetter "Windows"

    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        $SystemRoot = Get-SystemRootPath

        Write-Host "Configuring setup and adding script..." -ForegroundColor Green
        Set-ItemProperty-Logged -Path "HKLM:\BROKENSYSTEM\Setup" -Name SetupType -Value 2 -Type DWord -Force
        Set-ItemProperty-Logged -Path "HKLM:\BROKENSYSTEM\Setup" -Name CmdLine -Value "cmd.exe /c c:\temp\installwinrmhttps.cmd" -Type String -Force

        $installwinrmhttps = @"
@echo off
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -File C:\temp\configwinrmhttps.ps1
reg add "HKLM\SYSTEM\Setup" /v SetupType /t REG_DWORD /d 0 /f
reg add "HKLM\SYSTEM\Setup" /v CmdLine /t REG_SZ /d "" /f
del /F C:\temp\configwinrmhttps.ps1 > NUL
del /F C:\temp\winrm_https.pfx > NUL
del /F C:\temp\installwinrmhttps.cmd > NUL
"@

        $scriptContent = @'
$pwd = ConvertTo-SecureString "REPLACE_WITH_PASSWORD" -AsPlainText -Force

# 1) Import machine certificate (LocalMachine\My) and capture thumbprint
$imported = Import-PfxCertificate -FilePath "C:\temp\winrm_https.pfx" -CertStoreLocation "Cert:\LocalMachine\My" -Password $pwd
$thumb = $imported.Thumbprint

# 2) Enable WinRM for remoting + CredSSP (server)
Enable-PSRemoting -Force
Enable-WSManCredSSP -Role Server -Force

# 3) Firewall rule for 5986 (HTTPS)
New-NetFirewallRule -DisplayName "Windows Remote Management (HTTPS-In)" -Direction Inbound -LocalPort 5986 -Protocol TCP -Action Allow -Program "System" | Out-Null

# 4) Create HTTPS listener (remove any existing HTTPS listener first)
Get-ChildItem -Path WSMan:\localhost\Listener | Where-Object { $_.Keys -match "Transport=HTTPS" } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
New-Item -Path WSMan:\localhost\Listener -Transport HTTPS -Address * -Hostname $env:COMPUTERNAME -CertificateThumbprint $thumb -Port 5986 -Force | Out-Null

# 5) Ensure WinRM service is Auto + started
Set-Service -Name WinRM -StartupType Automatic
Restart-Service -Name WinRM -Force
'@

        Ensure-GuestTempDir
        $guestName = (Get-ItemProperty "$SystemRoot\Control\ComputerName\ComputerName").ComputerName

        $cert = New-SelfSignedCertificate -DnsName $guestName -CertStoreLocation 'Cert:\LocalMachine\My' -KeyExportPolicy Exportable
        # Generate a cryptographically random one-time PFX transport password (never hardcoded)
        $rngBytes = [System.Security.Cryptography.RandomNumberGenerator]::GetBytes(24)
        $PfxPassword = [System.Convert]::ToBase64String($rngBytes) -replace '[^A-Za-z0-9]', 'x'
        $certpwd = ConvertTo-SecureString $PfxPassword -AsPlainText -Force

        $hostPfx = Join-Path $WinDriveLetter "Temp\winrm_https.pfx"
        Export-PfxCertificate -Cert $cert -FilePath $hostPfx -Password $certpwd | Out-Null
        #Clean up locally
        Remove-Item-Logged -Path "Cert:\LocalMachine\My\$($cert.Thumbprint)" -Force

        New-Item-Logged -Path "$WinDriveLetter\Temp" -Name "installwinrmhttps.cmd" -ItemType File -Value $installwinrmhttps.ToString().Trim() -Force
        New-Item-Logged -Path "$WinDriveLetter\Temp" -Name "configwinrmhttps.ps1" -ItemType File -Value $scriptContent.ToString().Trim().Replace("REPLACE_WITH_PASSWORD", $PfxPassword) -Force -RedactValue
    }
    catch {
        Write-Error "SetWinRMHTTPSEnabled failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }

    Write-Host "Now start the VM and wait for the script to configure WinRM over HTTPS. (DO NO CONNECT WITH ENHANCED SESSION)." -ForegroundColor Yellow
    Write-Host "VM will restart after completing the WinRM configuration." -ForegroundColor Yellow

    if ($VMIPAddress) {
        Write-Host "Connect with: Enter-PSSession -ConnectionUri 'https://$($VMIPAddress):5986' -Credential (Get-Credential) -SessionOption (New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck)" -ForegroundColor Yellow
    }
    else {
        Write-Host "Connect with: Enter-PSSession -ConnectionUri 'https://<VMIP>:5986' -Credential (Get-Credential) -SessionOption (New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck)" -ForegroundColor Yellow
    }
}

################################################################################
# New repair functions
################################################################################

function FixBootSector {
    Write-Host "Fixing MBR/VBR boot sector..." -ForegroundColor Yellow
    try {
        if ($script:VMGen -ne 1) {
            Write-Warning "Gen2 (UEFI/GPT) VMs do not use MBR/VBR. Use -FixBoot to rebuild the EFI BCD store instead."
            return
        }

        $WinDrive = $script:WinDriveLetter.TrimEnd('\')
        $SysDrive = $script:BootDriveLetter.TrimEnd('\')

        # bootrec.exe only exists in WinRE/WinPE - not available on a regular Hyper-V host.
        # Use bootsect.exe (ships with Windows ADK) if present, otherwise fall back to bcdboot
        # which also rewrites the VBR as a side effect of recreating the boot files.
        $bootsect = $null
        $adkPaths = @(
            "$env:ProgramFiles\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\BCDBoot\bootsect.exe",
            "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\BCDBoot\bootsect.exe",
            "$env:ProgramFiles\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\x86\BCDBoot\bootsect.exe"
        )
        foreach ($p in $adkPaths) {
            if (Test-Path $p) { $bootsect = $p; break }
        }
        if (-not $bootsect) {
            $inPath = Get-Command bootsect.exe -ErrorAction SilentlyContinue
            if ($inPath) { $bootsect = $inPath.Source }
        }

        if ($bootsect) {
            Write-Host "Using bootsect.exe: $bootsect" -ForegroundColor Cyan
            # /nt60 = Windows Vista+ compatible VBR; /mbr = also rewrite MBR
            Invoke-Logged -Description "bootsect /nt60 /mbr" -Details @{ SysDrive = $SysDrive; Tool = $bootsect } -ScriptBlock {
                & $bootsect /nt60 $SysDrive /mbr /force
            }
            Write-Host "MBR and VBR repaired via bootsect. If boot still fails, try -FixBoot." -ForegroundColor Green
        }
        else {
            Write-Warning "bootsect.exe not found. It is only available via the Windows ADK (not installed) or inside WinRE."
            Write-Warning "Falling back to bcdboot, which rewrites the VBR and rebuilds boot files (equivalent to /fixboot)."
            Write-Warning "For a full MBR repair (/fixmbr equivalent), install the Windows ADK and re-run -FixBootSector."
            $rebuildCmd = "bcdboot $WinDrive\Windows /s $SysDrive /v /f BIOS"
            Invoke-Logged -Description "bcdboot fallback (VBR repair)" -Details @{ Command = $rebuildCmd } -ScriptBlock {
                & cmd.exe /c $rebuildCmd
            }
            Write-Host "VBR repaired via bcdboot. BCD entries also rebuilt as part of this operation." -ForegroundColor Green
        }
    }
    catch {
        Write-Error "FixBootSector failed: $_"
        throw
    }
}

function DisableStartupRepair {
    Write-Host "Disabling automatic startup repair loop..." -ForegroundColor Yellow
    try {
        $storePath = Get-BcdStorePath -Generation $script:VMGen -BootDrive $script:BootDriveLetter
        $identifier = Get-BcdBootLoaderId -StorePath $storePath
        if (-not $identifier) { return }

        Invoke-BcdEdit -StorePath $storePath -Command "/set $identifier recoveryenabled No"
        Invoke-BcdEdit -StorePath $storePath -Command "/set $identifier bootstatuspolicy IgnoreAllFailures"
        Invoke-BcdEdit -StorePath $storePath -Command "/set {bootmgr} displaybootmenu no"
        Write-Host "Startup repair disabled. The VM will skip WinRE on next failed boot." -ForegroundColor Green
        Write-Host "To re-enable: use -EnableStartupRepair" -ForegroundColor DarkCyan
    }
    catch {
        Write-Error "DisableStartupRepair failed: $_"
        throw
    }
}

function EnableStartupRepair {
    Write-Host "Enabling automatic startup repair..." -ForegroundColor Yellow
    try {
        $storePath = Get-BcdStorePath -Generation $script:VMGen -BootDrive $script:BootDriveLetter
        $identifier = Get-BcdBootLoaderId -StorePath $storePath
        if (-not $identifier) { return }

        Invoke-BcdEdit -StorePath $storePath -Command "/set $identifier recoveryenabled Yes"
        Invoke-BcdEdit -StorePath $storePath -Command "/deletevalue $identifier bootstatuspolicy"
        Invoke-BcdEdit -StorePath $storePath -Command "/set {bootmgr} displaybootmenu no"
        Write-Host "Startup repair re-enabled. Windows will enter WinRE on boot failure." -ForegroundColor Green
    }
    catch {
        Write-Error "EnableStartupRepair failed: $_"
        throw
    }
}

function ResetLocalAdminPassword {
    Write-Host "Resetting local administrator account password via startup script..." -ForegroundColor Yellow
    Write-Host "Please provide the new credentials (username = existing local admin account, no .\ prefix):" -ForegroundColor Yellow
    $Credential = Get-Credential

    $Username = $Credential.GetNetworkCredential().UserName
    $Password = $Credential.GetNetworkCredential().Password

    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        Set-ItemProperty-Logged -Path "HKLM:\BROKENSYSTEM\Setup" -Name SetupType -Value 2 -Type DWord -Force
        Set-ItemProperty-Logged -Path "HKLM:\BROKENSYSTEM\Setup" -Name CmdLine -Value "cmd.exe /c c:\temp\resetpwd.cmd" -Type String -Force

        $resetScript = @"
@echo off
net user Username "Password" /Y
reg add "HKLM\SYSTEM\Setup" /v SetupType /t REG_DWORD /d 0 /f
reg add "HKLM\SYSTEM\Setup" /v CmdLine /t REG_SZ /d "" /f
del /F C:\Temp\resetpwd.cmd > NUL
"@
        Ensure-GuestTempDir
        New-Item-Logged -Path "$script:WinDriveLetter\Temp" -Name "resetpwd.cmd" -ItemType File -Value $resetScript.Trim().Replace("Username", $Username).Replace("Password", $Password) -Force -RedactValue

        Write-Host "Password reset script placed. The password for '$Username' will be changed on next VM boot." -ForegroundColor Green
    }
    catch {
        Write-Error "ResetLocalAdminPassword failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

function DisableThirdPartyDrivers {
    if (-not (Confirm-CriticalOperation -Operation 'Disable Third-Party Drivers (-DisableThirdPartyDrivers)' -Details @"
Sets Start=4 (Disabled) for all non-Microsoft Boot and System kernel drivers.
Revert commands are printed after completion, or use -EnableThirdPartyDrivers.
"@)) { return }

    Write-Host "Enumerating and disabling non-Microsoft kernel drivers..." -ForegroundColor Yellow
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        $SystemRoot   = Get-SystemRootPath
        $ServicesRoot = "$SystemRoot\Services"

        # Kernel-mode start types: 0=Boot, 1=System, 2=Auto, 3=Manual, 4=Disabled
        # Target Boot(0) and System(1) drivers that are third-party

        $disabled = @()
        Get-ChildItem $ServicesRoot -ErrorAction SilentlyContinue | ForEach-Object {
            $svcPath = $_.PSPath
            $props = Get-ItemProperty -Path $svcPath -ErrorAction SilentlyContinue
            # Only kernel/filesystem drivers (Type 1=kernel, 2=filesystem)
            if ($props.Type -notin @(1, 2)) { return }
            # Only Boot and System start drivers
            if ($props.Start -notin @(0, 1)) { return }
            # Skip if already disabled
            if ($props.Start -eq 4) { return }

            $imagePathRaw = $props.ImagePath
            if (-not $imagePathRaw) { return }
            $imagePath = $imagePathRaw `
                -replace '(?i)\\SystemRoot\\', "$script:WinDriveLetter\Windows\" `
                -replace '(?i)%SystemRoot%',   "$script:WinDriveLetter\Windows" `
                -replace '(?i)\\\?\?\\',   '' `
                -replace '(?i)^system32\\',   "$script:WinDriveLetter\Windows\System32\\"
            if ($imagePath -match '^(.+?\.(?:sys|exe))') { $imagePath = $Matches[1] }

            # Read the binary to check the company name in version info.
            # If the binary is missing, treat the driver as third-party — a missing
            # Boot/System binary will cause a BSOD on next boot regardless of vendor.
            $isMicrosoft   = $false
            $binaryMissing = $false
            if (Test-Path $imagePath) {
                $vi = (Get-Item $imagePath -ErrorAction SilentlyContinue).VersionInfo
                if ($vi -and $vi.CompanyName -match 'Microsoft') { $isMicrosoft = $true }
            }
            else {
                $binaryMissing = $true   # missing binary -> will BSOD -> must disable
            }

            if (-not $isMicrosoft) {
                $svcName = $_.PSChildName
                $tag = if ($binaryMissing) { ' [binary missing - will BSOD on boot]' } else { '' }
                Write-Host "  Disabling: $svcName$tag  ($imagePathRaw)" -ForegroundColor Yellow
                Set-ItemProperty-Logged -Path $svcPath -Name Start -Value 4 -Type DWord -Force
                $disabled += [PSCustomObject]@{ Service = $svcName; ImagePath = $imagePathRaw; PreviousStart = $props.Start }
            }
        }

        if ($disabled.Count -eq 0) {
            Write-Host "No non-Microsoft Boot/System drivers found to disable." -ForegroundColor Green
        }
        else {
            Write-Host "`nDisabled $($disabled.Count) third-party driver(s)." -ForegroundColor Green
            Write-Host "`n--- TO RE-ENABLE (run -EnableThirdPartyDrivers or run on the VM after recovery) ---" -ForegroundColor Cyan
            foreach ($d in $disabled) {
                $livePath = "HKLM:\SYSTEM\CurrentControlSet\Services\$($d.Service)"
                Write-Host "Set-ItemProperty -Path '$livePath' -Name Start -Value $($d.PreviousStart) -Type DWord -Force" -ForegroundColor White
            }
            Write-Host "----------------------------------------------------`n" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Error "DisableThirdPartyDrivers failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

function GetServicesReport {
    param(
        [switch]$IncludeServices,
        [switch]$IssuesOnly
    )
    # Enumerates every entry under Services in the active ControlSet and produces a
    # grouped, colour-coded report. Each row shows:
    #   Name | [Pres] | ErCtl | Vendor | ImagePath
    # By default only kernel/filesystem drivers are shown (Type 1/2/4/8).
    # Pass -IncludeServices to include Win32 services as well.
    # Pass -IssuesOnly to suppress healthy Microsoft rows and display only:
    #   - Missing binaries
    #   - Non-Microsoft vendors
    #   - ErrorControl >= 2 (Severe/Critical)
    # Rows are sorted by Start value (Boot -> System -> Automatic -> Manual -> Disabled)
    # then alphabetically within each group.

    $reportLabel = if ($IncludeServices) { 'services/drivers' } else { 'drivers' }
    $modeLabel   = if ($IssuesOnly) { ' (issues only)' } else { '' }
    Write-Host "`nBuilding offline $reportLabel report$modeLabel..." -ForegroundColor Cyan
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        $SystemRoot   = Get-SystemRootPath
        $ServicesRoot = "$SystemRoot\Services"

        $startNames = @{ 0='Boot'; 1='System'; 2='Automatic'; 3='Manual'; 4='Disabled' }
        $typeNames  = @{ 1='KernelDriver'; 2='FileSystemDriver'; 4='Adapter'; 8='Recognizer';
                         16='Win32Own'; 32='Win32Share'; 256='Interactive' }

        $rows = [System.Collections.Generic.List[PSCustomObject]]::new()

        Get-ChildItem $ServicesRoot -ErrorAction SilentlyContinue | ForEach-Object {
            $props = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
            if (-not $props) { return }

            $startVal = $props.Start
            $typeVal  = $props.Type
            # Skip entries with no Start or no Type (e.g. sub-keys that are not real services)
            if ($null -eq $startVal -or $null -eq $typeVal) { return }

            $startLabel = if ($startNames.ContainsKey([int]$startVal)) { $startNames[[int]$startVal] } else { "Unknown($startVal)" }
            $typeLabel  = if ($typeNames.ContainsKey([int]$typeVal))   { $typeNames[[int]$typeVal]   } else { "Type($typeVal)" }

            $imgRaw  = $props.ImagePath
            $imgPath = $null
            $present = $null    # $null = no ImagePath registered
            $vendor  = 'N/A'

            if ($imgRaw) {
                $winDrive = $script:WinDriveLetter.TrimEnd('\')
                $imgPath = $imgRaw `
                    -replace '(?i)\\SystemRoot\\',  "$winDrive\Windows\" `
                    -replace '(?i)%SystemRoot%',    "$winDrive\Windows" `
                    -replace '(?i)\\\?\?\\',       '' `
                    -replace '(?i)^system32\\',    "$winDrive\Windows\System32\" `
                    -replace '(?i)^"?[A-Z]:\\',    "$winDrive\"   # rebase any absolute guest path (C:\...) to the offline drive
                if ($imgPath -match '^(.+?\.(?:sys|exe|dll))') { $imgPath = $Matches[1] }

                if (Test-Path $imgPath) {
                    $present = $true
                    $vi = (Get-Item $imgPath -ErrorAction SilentlyContinue).VersionInfo
                    $vendor = if ($vi -and $vi.CompanyName) { $vi.CompanyName.Trim() } else { '(no version info)' }
                } else {
                    $present = $false
                    $vendor  = '(binary missing)'
                }
            }

            $rows.Add([PSCustomObject]@{
                Name          = $_.PSChildName
                Start         = $startLabel
                StartVal      = [int]$startVal
                TypeVal       = [int]$typeVal
                Type          = $typeLabel
                Group         = if ($props.Group) { $props.Group } else { '' }
                Vendor        = $vendor
                BinaryPresent = $present
                ImagePath     = if ($imgRaw) { $imgRaw } else { '' }
                ErrorControl  = if ($null -ne $props.ErrorControl) { [int]$props.ErrorControl } else { $null }
            })
        }

        # Filter to drivers only unless -IncludeServices was given.
        # Driver types: 1=KernelDriver, 2=FileSystemDriver, 4=Adapter, 8=Recognizer
        if (-not $IncludeServices) {
            $rows = [System.Collections.Generic.List[PSCustomObject]]($rows | Where-Object { $_.TypeVal -in @(1,2,4,8) })
        }

        # Filter to issue rows only when -IssuesOnly is given.
        # An issue row is one that warrants attention:
        #   - Missing binary (any vendor)
        #   - Non-Microsoft vendor with a binary present
        #   - ErrorControl >= 2 (Severe/Critical) AND non-Microsoft vendor
        #     (Microsoft inbox drivers legitimately carry CRIT/Sev ErrorControl - that is normal)
        if ($IssuesOnly) {
            $rows = [System.Collections.Generic.List[PSCustomObject]]($rows | Where-Object {
                $isMsVendor = $_.Vendor -match 'Microsoft'
                $_.BinaryPresent -eq $false -or
                ($_.BinaryPresent -eq $true -and -not $isMsVendor) -or
                ($null -ne $_.ErrorControl -and $_.ErrorControl -ge 2 -and -not $isMsVendor)
            })
        }

        if ($rows.Count -eq 0) {
            $noneMsg = if ($IssuesOnly) { "No issues found in the offline $reportLabel ControlSet - all entries look healthy." } else { "No $reportLabel entries found in the offline ControlSet." }
            Write-Host "  $noneMsg" -ForegroundColor $(if ($IssuesOnly) { 'Green' } else { 'DarkGray' })
            return
        }

        # Group by Start value, then sort alphabetically within each group
        $grouped = $rows | Sort-Object StartVal, Name

        $currentStart = $null
        foreach ($r in $grouped) {
            if ($r.Start -ne $currentStart) {
                $currentStart = $r.Start
                $headerColor  = switch ($r.Start) {
                    'Boot'      { 'Red'    }
                    'System'    { 'Yellow' }
                    'Automatic' { 'Cyan'   }
                    'Manual'    { 'White'  }
                    'Disabled'  { 'DarkGray' }
                    default     { 'Gray'   }
                }
                Write-Host "`n  == $($r.Start.ToUpper()) =========================================" -ForegroundColor $headerColor
                Write-Host ("  {0,-30} {1,-28} {2,-6} {3,-6} {4}" -f 'Name','Vendor','[Pres]','ErCtl','ImagePath') -ForegroundColor DarkGray
                Write-Host ("  {0,-30} {1,-28} {2,-6} {3,-6} {4}" -f ('-'*30),('-'*28),'------','------',('-'*40)) -ForegroundColor DarkGray
            }

            # Colour logic: missing binary = Red, non-Microsoft = Yellow, Microsoft = Green, N/A = Gray
            $isMicrosoft = $r.Vendor -match 'Microsoft'
            $isMissing   = $r.BinaryPresent -eq $false
            $noPath      = $null -eq $r.BinaryPresent

            $rowColor = if ($isMissing)      { 'Red'     }
                        elseif ($noPath)     { 'DarkGray'}
                        elseif ($isMicrosoft){ 'Green'   }
                        else                 { 'Yellow'  }

            $presTag = if ($isMissing) { 'MISS' } elseif ($noPath) { ' -- ' } else { ' OK ' }
            $vendorShort = if ($r.Vendor.Length -gt 28) { $r.Vendor.Substring(0,25) + '...' } else { $r.Vendor }

            # ErrorControl: 0=Ignore, 1=Normal, 2=Severe (LKGC fallback), 3=Critical (boot failure)
            $errCtlLabel = if ($null -eq $r.ErrorControl) { '    ' } else {
                switch ($r.ErrorControl) { 0{'Ign'} 1{'Norm'} 2{'Sev!'} 3{'CRIT'} default{"EC$($r.ErrorControl)"} }
            }

            Write-Host ("  {0,-30} {1,-28} [{2}] {3,-6} {4}" -f $r.Name, $vendorShort, $presTag, $errCtlLabel, $r.ImagePath) -ForegroundColor $rowColor
        }

        # Summary
        $total    = $rows.Count
        $missing  = @($rows | Where-Object { $_.BinaryPresent -eq $false }).Count
        $nonMS    = @($rows | Where-Object { $_.BinaryPresent -eq $true -and $_.Vendor -notmatch 'Microsoft' }).Count
        $boot_sys = @($rows | Where-Object { $_.StartVal -in @(0,1) }).Count
        $sevCrit  = @($rows | Where-Object { $null -ne $_.ErrorControl -and $_.ErrorControl -ge 2 }).Count

        Write-Host "`n  ===========================================================" -ForegroundColor Cyan
        Write-Host ("  Total {0}: {1}  |  Boot/System: {2}  |  Missing binary: {3}  |  Non-Microsoft: {4}  |  Severe/Critical EC: {5}" `
            -f $reportLabel, $total, $boot_sys, $missing, $nonMS, $sevCrit) -ForegroundColor Cyan
        if ($missing -gt 0) {
            Write-Host "  [!] Missing-binary Boot/System drivers will cause a BSOD - run -DisableThirdPartyDrivers to neutralise them." -ForegroundColor Red
        }
        if ($nonMS -gt 0) {
            Write-Host "  [!] Non-Microsoft drivers present - verify each is expected for this guest OS." -ForegroundColor Yellow
        }
        if ($sevCrit -gt 0) {
            Write-Host "  [!] $sevCrit driver(s) with ErrorControl Sev!(2) or CRIT(3) - failure of these at boot will trigger LKGC fallback or halt the system." -ForegroundColor Yellow
        }
        Write-Host "  ===========================================================`n" -ForegroundColor Cyan
    }
    catch {
        Write-Error "GetServicesReport failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

function EnableThirdPartyDrivers {
    Write-Host "Re-enabling previously disabled third-party kernel drivers..." -ForegroundColor Yellow
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        $SystemRoot   = Get-SystemRootPath
        $ServicesRoot = "$SystemRoot\Services"

        $reEnabled = @()
        Get-ChildItem $ServicesRoot -ErrorAction SilentlyContinue | ForEach-Object {
            $svcPath = $_.PSPath
            $props = Get-ItemProperty -Path $svcPath -ErrorAction SilentlyContinue
            # Only disabled kernel/filesystem drivers
            if ($props.Type -notin @(1, 2)) { return }
            if ($props.Start -ne 4) { return }

            $imagePathRaw = $props.ImagePath
            if (-not $imagePathRaw) { return }
            $imagePath = $imagePathRaw `
                -replace '(?i)\\SystemRoot\\', "$script:WinDriveLetter\Windows\" `
                -replace '(?i)%SystemRoot%',   "$script:WinDriveLetter\Windows" `
                -replace '(?i)\\\?\?\\',   '' `
                -replace '(?i)^system32\\',   "$script:WinDriveLetter\Windows\System32\\"
            if ($imagePath -match '^(.+?\.(?:sys|exe))') { $imagePath = $Matches[1] }

            $isMicrosoft = $false
            if (Test-Path $imagePath) {
                $vi = (Get-Item $imagePath -ErrorAction SilentlyContinue).VersionInfo
                if ($vi -and $vi.CompanyName -match 'Microsoft') { $isMicrosoft = $true }
            }
            else {
                # Binary is missing - warn and skip re-enabling (re-enabling a driver with a
                # missing binary will cause a BSOD on boot).
                Write-Warning "Skipping $($_.PSChildName) - binary is missing ($imagePath). Restore the driver file before re-enabling."
                return
            }

            if (-not $isMicrosoft) {
                $svcName = $_.PSChildName
                # Restore to System (1) start — safest default for a previously Boot/System driver
                Write-Host "  Re-enabling: $svcName" -ForegroundColor Yellow
                Set-ItemProperty-Logged -Path $svcPath -Name Start -Value 1 -Type DWord -Force
                $reEnabled += $svcName
            }
        }

        if ($reEnabled.Count -eq 0) {
            Write-Host "No disabled third-party Boot/System drivers found to re-enable." -ForegroundColor Green
        }
        else {
            Write-Host "`nRe-enabled $($reEnabled.Count) driver(s): $($reEnabled -join ', ')" -ForegroundColor Green
            Write-Warning "Start value restored to 1 (System). If a driver was originally Boot (0), adjust manually after verifying VM boots." 
        }
    }
    catch {
        Write-Error "EnableThirdPartyDrivers failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

# Remove the Credential Guard UEFI firmware variable that persists when CG is enabled with UEFI lock.
# Registry changes alone (zeroing LsaCfgFlags) are insufficient because the EFI NVRAM variable
# survives reboots independently of the registry.
# This script repairs offline disks that are later reattached to their Azure VM, so the only
# viable offline method is injecting a SecConfig.efi one-time BCD boot entry: on the VM's next
# boot (on Azure) the EFI application runs and clears the UEFI variable inside the guest firmware.
# NOTE: -FixBoot (RebuildBCD) only recreates the BCD store and does NOT clear this EFI variable.
function Remove-CredentialGuardUefiLock {
    # ------------------------------------------------------------------
    # bcdedit + SecConfig.efi one-time-boot entry
    # Copies the Microsoft-supplied SecConfig.efi tool to the guest's EFI
    # partition on the offline disk and adds a one-time BCD bootsequence
    # entry. When the disk is reattached to the Azure VM and it boots,
    # SecConfig.efi runs inside the VM and clears the UEFI NVRAM variable.
    # Azure Gen2/Trusted Launch VMs auto-accept the confirmation (no
    # physical key press required).
    # ------------------------------------------------------------------
    if ($script:VMGen -ne 2 -or [string]::IsNullOrWhiteSpace($script:BootDriveLetter)) {
        Write-Warning "  SecConfig.efi method requires a mounted Gen2/UEFI EFI partition - skipping."
        Write-Error "UEFI lock could not be cleared: EFI partition not available."
        return
    }

    $efiVolume  = $script:BootDriveLetter.TrimEnd('\')
    $storePath  = "$efiVolume\EFI\Microsoft\Boot\BCD"
    $secDest    = "$efiVolume\EFI\Microsoft\Boot\SecConfig.efi"
    $secSrc     = "$env:SystemRoot\System32\SecConfig.efi"

    if (-not (Test-Path $secSrc)) {
        Write-Warning "  SecConfig.efi not found at '$secSrc' on this host."
        Write-Error "UEFI lock could not be cleared: SecConfig.efi unavailable."
        return
    }

    if (-not (Test-Path $storePath)) {
        Write-Warning "  Offline BCD not found at '$storePath'."
        Write-Error "UEFI lock could not be cleared: offline BCD not found."
        return
    }

    try {
        Write-Host "  Configuring SecConfig.efi one-time boot entry in offline BCD..." -ForegroundColor Cyan
        Copy-Item-Logged -Path $secSrc -Destination $secDest -Force

        # Well-known GUID defined by Microsoft for the CG disable tool boot entry
        $cgGuid = '{0cb3b571-2f2e-4343-a879-d86a476d7215}'

        # Remove any pre-existing entry with this GUID to ensure a clean state
        $deleteBcdCmd = "bcdedit /store `"$storePath`" /delete $cgGuid /cleanup"
        Write-Host "  [exec] $deleteBcdCmd" -ForegroundColor DarkGray
        & cmd.exe /c $deleteBcdCmd 2>&1 | Out-Null

        # Create the one-time boot sequence entry
        $bcdCmds = @(
            "/create $cgGuid /d `"CredentialGuard UEFI Disable Tool`" /application osloader",
            "/set $cgGuid path `"\EFI\Microsoft\Boot\SecConfig.efi`"",
            "/set {bootmgr} bootsequence $cgGuid",
            "/set $cgGuid loadoptions DISABLE-LSA-ISO",
            "/set $cgGuid device partition=${efiVolume}:"
        )
        foreach ($cmd in $bcdCmds) {
            Invoke-BcdEdit -StorePath $storePath -Command $cmd
        }

        Write-Host "  [OK] SecConfig.efi one-time boot entry added to offline BCD." -ForegroundColor Green
        Write-Host ("  On the VM's first boot back on Azure, SecConfig.efi will run inside the VM " +
                    "and clear the Credential Guard EFI NVRAM variable automatically.") -ForegroundColor Cyan
    }
    catch {
        Write-Warning "  bcdedit SecConfig.efi setup failed: $_"
        Write-Error "UEFI lock could not be cleared."
    }
}

function DisableCredentialGuard {
    Write-Host "Disabling Credential Guard and LSA protection..." -ForegroundColor Yellow
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    $uefiLockWasActive  = $false
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SOFTWARE"
    try {
        $SystemRoot = Get-SystemRootPath

        $lsaPath   = "$SystemRoot\Control\Lsa"
        $cgPath    = "HKLM:\BROKENSOFTWARE\Policies\Microsoft\Windows\DeviceGuard"
        $cgSysPath = "$SystemRoot\Control\DeviceGuard"

        # Snapshot for restore
        $beforeRunAsPPL  = (Get-ItemProperty $lsaPath   -ErrorAction SilentlyContinue).RunAsPPL
        $beforeLsaCfg    = (Get-ItemProperty $cgSysPath -ErrorAction SilentlyContinue).LsaCfgFlags
        $beforeCGEnabled = (Get-ItemProperty $cgPath    -ErrorAction SilentlyContinue).EnableVirtualizationBasedSecurity

        # Detect UEFI lock: LsaCfgFlags=1 in Control\Lsa means CG was enabled WITH UEFI lock.
        # The EFI NVRAM variable survives registry changes and must be cleared separately.
        # LsaCfgFlags=2 means enabled WITHOUT lock — registry zeroing is sufficient.
        $lsaLsaCfgFlags = (Get-ItemProperty $lsaPath -ErrorAction SilentlyContinue).LsaCfgFlags
        if ($lsaLsaCfgFlags -eq 1) {
            $uefiLockWasActive = $true
            Write-Warning ("Credential Guard was enabled WITH UEFI lock (Control\Lsa\LsaCfgFlags=1). " +
                           "An EFI NVRAM variable is also set - registry changes alone are not sufficient. " +
                           "The UEFI firmware variable will be cleared after applying registry changes.")
        }
        elseif ($lsaLsaCfgFlags -eq 2) {
            Write-Host "  Credential Guard mode: WITHOUT UEFI lock - registry changes are sufficient." -ForegroundColor Cyan
        }

        # Zero all registry knobs (both paths that can enforce CG)
        Set-ItemProperty-Logged -Path $lsaPath   -Name RunAsPPL    -Value 0 -Type DWord -Force
        Set-ItemProperty-Logged -Path $lsaPath   -Name LsaCfgFlags -Value 0 -Type DWord -Force
        Set-ItemProperty-Logged -Path $cgSysPath -Name LsaCfgFlags -Value 0 -Type DWord -Force
        if (Test-Path $cgPath) {
            Set-ItemProperty-Logged -Path $cgPath -Name EnableVirtualizationBasedSecurity -Value 0 -Type DWord -Force
            Set-ItemProperty-Logged -Path $cgPath -Name LsaCfgFlags                       -Value 0 -Type DWord -Force
        }

        Write-Host "Credential Guard registry settings cleared. Use -EnableCredentialGuard to revert." -ForegroundColor Green
        Write-Host "`n--- REVERT COMMANDS (run on the VM after recovery) ---" -ForegroundColor Cyan
        $lsaLive   = "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa"
        $cgSysLive = "HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard"
        $cgLive    = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard"
        Write-Host "Set-ItemProperty -Path '$lsaLive'   -Name RunAsPPL    -Value $(if ($null -ne $beforeRunAsPPL) { $beforeRunAsPPL } else { '1 # (was not set)' }) -Type DWord -Force" -ForegroundColor White
        Write-Host "Set-ItemProperty -Path '$cgSysLive' -Name LsaCfgFlags -Value $(if ($null -ne $beforeLsaCfg)  { $beforeLsaCfg  } else { '1 # (was not set)' }) -Type DWord -Force" -ForegroundColor White
        if ($null -ne $beforeCGEnabled) {
            Write-Host "Set-ItemProperty -Path '$cgLive' -Name EnableVirtualizationBasedSecurity -Value $beforeCGEnabled -Type DWord -Force" -ForegroundColor White
        }
        Write-Host "------------------------------------------------------`n" -ForegroundColor Cyan
    }
    catch {
        Write-Error "DisableCredentialGuard failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
        UnmountOffHive -Hive "SOFTWARE"
    }

    # Hives are now safely unmounted. Clear the EFI NVRAM variable if UEFI lock was active.
    if ($uefiLockWasActive) {
        Write-Host "`nClearing Credential Guard EFI NVRAM variable (UEFI lock)..." -ForegroundColor Yellow
        Remove-CredentialGuardUefiLock
    }
}

function EnableCredentialGuard {
    Write-Host "Re-enabling Credential Guard and LSA protection..." -ForegroundColor Yellow

    # ----------------------------------------------------------------
    # Safety gate 1: Credential Guard requires UEFI Secure Boot (Gen2/GPT).
    # Writing the registry flags on a Gen1/MBR VM causes a no-boot scenario
    # because Windows cannot satisfy the Secure Boot requirement at startup.
    # $script:VMGen is already set by Repair-OfflineDisk via Get-DiskGeneration
    # (GPT -> 2, MBR -> 1), so this check is reliable even when -DiskNumber
    # is used instead of -VMName.
    # ----------------------------------------------------------------
    if ($script:VMGen -eq 1) {
        Write-Error ("BLOCKED: This disk uses MBR/BIOS (Gen1). Credential Guard requires UEFI Secure Boot " +
                     "(Gen2/GPT). Enabling it on a Gen1 VM writes flags that prevent Windows from booting. Aborting.")
        return
    }
    if ($null -eq $script:VMGen) {
        Write-Error ("BLOCKED: Disk partition style could not be determined (expected MBR=Gen1 or GPT=Gen2). " +
                     "Credential Guard requires UEFI/GPT (Gen2). Aborting to prevent a potential no-boot scenario.")
        return
    }

    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SOFTWARE"
    try {
        $SystemRoot = Get-SystemRootPath

        # ----------------------------------------------------------------
        # Safety gate 2: Credential Guard is only supported on Enterprise,
        # Education, and Server editions. Enabling it on Home/Pro/Core writes
        # the same registry flags but Windows will fail a boot-time licence
        # check resulting in an unbootable VM.
        # ----------------------------------------------------------------
        $winVer      = Get-ItemProperty "HKLM:\BROKENSOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue
        $editionId   = $winVer.EditionID
        $productName = $winVer.ProductName
        $unsupportedEditions = @('Home', 'Pro', 'CoreSingleLanguage', 'Core', 'ProWorkstation', 'ProfessionalWorkstation')
        $editionBlocked = $unsupportedEditions | Where-Object { $editionId -match $_ }
        if ($editionBlocked) {
            Write-Error ("BLOCKED: Windows edition '$editionId' ($productName) does not support Credential Guard. " +
                         "Only Enterprise, Education, and Server editions are supported. Aborting.")
            return
        }
        if ($editionId) {
            Write-Host "  Windows edition: $editionId ($productName) - supported." -ForegroundColor Cyan
        } else {
            Write-Warning "  Could not read Windows edition from offline hive; proceeding with caution."
        }

        $lsaPath   = "$SystemRoot\Control\Lsa"
        $cgSysPath = "$SystemRoot\Control\DeviceGuard"
        $cgPath    = "HKLM:\BROKENSOFTWARE\Policies\Microsoft\Windows\DeviceGuard"

        Set-ItemProperty-Logged -Path $lsaPath   -Name RunAsPPL     -Value 1 -Type DWord -Force
        Set-ItemProperty-Logged -Path $cgSysPath -Name LsaCfgFlags  -Value 1 -Type DWord -Force
        if (Test-Path $cgPath) {
            Set-ItemProperty-Logged -Path $cgPath -Name EnableVirtualizationBasedSecurity -Value 1 -Type DWord -Force
            Set-ItemProperty-Logged -Path $cgPath -Name LsaCfgFlags                       -Value 1 -Type DWord -Force
        }

        Write-Host "Credential Guard and LSA protection re-enabled." -ForegroundColor Green
        Write-Warning ("Before starting the VM, verify in Hyper-V settings that: " +
                       "(1) Secure Boot is enabled, (2) Virtualization Based Security (VBS) is enabled. " +
                       "If either is missing the VM will fail to boot. Use -DisableCredentialGuard to revert.")
    }
    catch {
        Write-Error "EnableCredentialGuard failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
        UnmountOffHive -Hive "SOFTWARE"
    }
}

function DisableAppLocker {
    # AppLocker can prevent Windows from booting when an Enforce-mode policy incorrectly
    # blocks a system binary, service DLL, or startup script. The fix has two parts:
    #
    #  1. Set all AppLocker rule collection EnforcementMode values to 0 (Not Configured)
    #     in both the GPO-policy path (SOFTWARE\Policies) and the local policy path.
    #     This is non-destructive: the rules are preserved but not enforced, so the
    #     admin can re-enable them once the VM is healthy.
    #
    #  2. Disable the Application Identity service (AppIDSvc). AppLocker relies on this
    #     service at runtime; disabling it prevents enforcement even if the policy paths
    #     are re-applied by GPO refresh before an admin can investigate.
    #
    # EnforcementMode values:
    #   0 = Not Configured (rules ignored)
    #   1 = Enforce rules
    #   2 = Audit only

    Write-Host "Disabling AppLocker enforcement..." -ForegroundColor Yellow
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SOFTWARE"
    try {
        $SystemRoot = Get-SystemRootPath

        # Rule collection names that AppLocker enforces
        $ruleCollections = @('Exe', 'Dll', 'Script', 'Msi', 'Appx')

        # Two registry paths can carry AppLocker policy:
        #   - GPO/MDM-applied:  SOFTWARE\Policies\Microsoft\Windows\SrpV2
        #   - Local policy:     SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\AppV (rare)
        # The GPO path is by far the most common source of broken policies on Azure VMs.
        $srpPaths = @(
            'HKLM:\BROKENSOFTWARE\Policies\Microsoft\Windows\SrpV2',
            'HKLM:\BROKENSOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\AppV\SrpV2'
        )

        # --- Pass 1: scan for any actively-enforced collections (mode 1 or 2) ---
        $enforced = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($srpBase in $srpPaths) {
            if (-not (Test-Path $srpBase)) { continue }
            foreach ($collection in $ruleCollections) {
                $collPath = "$srpBase\$collection"
                if (-not (Test-Path $collPath)) { continue }
                $current_mode = (Get-ItemProperty $collPath -ErrorAction SilentlyContinue).EnforcementMode
                if ($null -ne $current_mode -and $current_mode -ne 0) {
                    $modeLabel = if ($current_mode -eq 1) { 'Enforce' } else { 'Audit' }
                    Write-Host "  [$collection] EnforcementMode = $current_mode ($modeLabel) - will disable." -ForegroundColor Yellow
                    $enforced.Add(@{ Path = $collPath; Collection = $collection; Mode = $current_mode })
                }
            }
        }

        # --- Check AppIDSvc start value ---
        $appIdPath    = "$SystemRoot\Services\AppIDSvc"
        $appIdStart   = if (Test-Path $appIdPath) { (Get-ItemProperty $appIdPath -ErrorAction SilentlyContinue).Start } else { $null }
        # Service runs if Start is 2 (Auto) or 3 (Manual/demand); 4 = Disabled, 1 = Boot (unlikely)
        $appIdRunning = $null -ne $appIdStart -and $appIdStart -in @(1, 2, 3)

        # --- Early exit if nothing is active ---
        if ($enforced.Count -eq 0 -and -not $appIdRunning) {
            Write-Host "  AppLocker is not enforced (no active rule collections, AppIDSvc not running). No changes made." -ForegroundColor Green
            return
        }

        # --- Pass 2: disable only what is active ---
        foreach ($item in $enforced) {
            Write-Host "  [$($item.Collection)] EnforcementMode: $($item.Mode) -> 0 (Not Configured)" -ForegroundColor Cyan
            Set-ItemProperty-Logged -Path $item.Path -Name EnforcementMode -Value 0 -Type DWord -Force
        }

        # Disable Application Identity service so AppLocker cannot enforce at runtime
        # even if GPO refreshes the policy keys before an admin can intervene.
        if ($appIdRunning) {
            Write-Host "  AppIDSvc Start: $appIdStart -> 4 (Disabled)" -ForegroundColor Cyan
            Set-ItemProperty-Logged -Path $appIdPath -Name Start -Value 4 -Type DWord -Force
        } elseif (-not (Test-Path $appIdPath)) {
            Write-Host "  AppIDSvc registry key not found - service not installed on this image." -ForegroundColor DarkGray
        }

        Write-Host "AppLocker enforcement disabled." -ForegroundColor Green
        Write-Host ("  NOTE: AppLocker rules are preserved (EnforcementMode=0). " +
                    "To re-enable, set EnforcementMode=1 under SOFTWARE\Policies\Microsoft\Windows\SrpV2\<Collection> " +
                    "and set AppIDSvc Start back to 2 (Automatic) on the live VM.") -ForegroundColor Cyan
        Write-Warning ("If AppLocker policy is delivered by Intune/MDM, the offline registry path will be empty " +
                       "and policy will reapply on next MDM sync. Resolve via Intune before starting the VM.")
    }
    catch {
        Write-Error "DisableAppLocker failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
        UnmountOffHive -Hive "SOFTWARE"
    }
}

function FixSanPolicy {
    Write-Host "Setting SAN policy to OnlineAll so all disks come online on boot..." -ForegroundColor Yellow
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        $SystemRoot = Get-SystemRootPath
        $sanPath    = "$SystemRoot\Services\partmgr\Parameters"

        $before = (Get-ItemProperty $sanPath -ErrorAction SilentlyContinue).SanPolicy
        Set-ItemProperty-Logged -Path $sanPath -Name SanPolicy -Value 1 -Type DWord -Force  # 1 = OnlineAll

        Write-Host "SAN policy set to OnlineAll (1). Previous value: $(if ($null -ne $before) { $before } else { 'not set' })." -ForegroundColor Green
        Write-Host "Revert: Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\partmgr\Parameters' -Name SanPolicy -Value $(if ($null -ne $before) { $before } else { 4 }) -Type DWord -Force" -ForegroundColor DarkCyan
    }
    catch {
        Write-Error "FixSanPolicy failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

function FixAzureGuestAgent {
    Write-Host "Enabling Azure Guest Agent and RdAgent services..." -ForegroundColor Yellow
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        $SystemRoot = Get-SystemRootPath

        $agents = @(
            @{ Name = 'WindowsAzureGuestAgent'; StartValue = 2 },
            @{ Name = 'RdAgent';                StartValue = 2 },
            @{ Name = 'WindowsAzureTelemetryService'; StartValue = 3 }
        )

        foreach ($agent in $agents) {
            $svcPath = "$SystemRoot\Services\$($agent.Name)"
            if (Test-Path $svcPath) {
                $before = (Get-ItemProperty $svcPath -ErrorAction SilentlyContinue).Start
                Set-ItemProperty-Logged -Path $svcPath -Name Start -Value $agent.StartValue -Type DWord -Force
                Write-Host "  $($agent.Name): Start set to $($agent.StartValue) (was $before)" -ForegroundColor Green
            }
            else {
                Write-Warning "  $($agent.Name) not found in registry - agent may not be installed."
            }
        }
        Write-Host "Azure Guest Agent services configured. Boot the VM for them to start." -ForegroundColor Green
    }
    catch {
        Write-Error "FixAzureGuestAgent failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

function InstallAzureVMAgentOffline {
    # Implements the offline VM Agent installation procedure from:
    # https://learn.microsoft.com/en-us/troubleshoot/azure/virtual-machines/windows/install-vm-agent-offline
    #
    # Source: this Hyper-V rescue VM (must have Azure VM Agent installed locally).
    # Steps:
    #   1. Validate the host has the required service registry keys and GuestAgent files.
    #   2. Backup any existing WindowsAzure folder on the offline disk.
    #   3. Copy GuestAgent_* folder(s) from C:\WindowsAzure on the host to the offline disk.
    #   4. Mount the offline SYSTEM hive and copy the service registry keys into it.
    #   5. Set service Start values to Auto so the agent starts on next boot.

    Write-Host "Installing Azure VM Agent in offline mode..." -ForegroundColor Yellow

    # ------------------------------------------------------------------
    # Step 1: Validate host sources
    # ------------------------------------------------------------------
    $requiredSvcs = @('WindowsAzureGuestAgent', 'RdAgent')
    foreach ($svc in $requiredSvcs) {
        $hostKey = "HKLM:\SYSTEM\ControlSet001\Services\$svc"
        if (-not (Test-Path $hostKey)) {
            Write-Error ("SKIPPED: Host registry key '$hostKey' not found. " +
                         "Azure VM Agent must be installed on this Hyper-V rescue VM to use -InstallAzureVMAgent. " +
                         "Download and install the agent from https://aka.ms/vmagentwin first.")
            return
        }
    }

    $hostWaFolder = 'C:\WindowsAzure'
    if (-not (Test-Path $hostWaFolder)) {
        Write-Error ("SKIPPED: '$hostWaFolder' not found on this host. " +
                     "Azure VM Agent must be installed on this Hyper-V rescue VM to use -InstallAzureVMAgent.")
        return
    }

    $agentFolders = Get-ChildItem -Path $hostWaFolder -Directory -Filter 'GuestAgent_*' -ErrorAction SilentlyContinue |
                    Sort-Object Name
    if (-not $agentFolders) {
        # Fallback: check ImagePath in the host registry to find the folder
        $imgPath = (Get-ItemProperty 'HKLM:\SYSTEM\ControlSet001\Services\WindowsAzureGuestAgent' `
                        -ErrorAction SilentlyContinue).ImagePath
        $imgDir  = if ($imgPath) { Split-Path $imgPath -Parent } else { $null }
        if ($imgDir -and (Test-Path $imgDir)) {
            Write-Host "  No GuestAgent_* folder found; using ImagePath directory: $imgDir" -ForegroundColor Cyan
            $agentFolders = @(Get-Item $imgDir)
        } else {
            Write-Error ("SKIPPED: No GuestAgent_* folders found under '$hostWaFolder' and ImagePath resolution failed. " +
                         "Cannot copy agent binaries to the offline disk.")
            return
        }
    }

    # ------------------------------------------------------------------
    # Step 2: Backup existing WindowsAzure folder on the offline disk
    # ------------------------------------------------------------------
    $offlineWaRoot = Join-Path $script:WinDriveLetter 'WindowsAzure'
    if (Test-Path $offlineWaRoot) {
        $bakPath = New-UniqueBackupPath -BasePath $offlineWaRoot -BakSuffix '.old'
        Write-Host "  Renaming existing WindowsAzure folder -> $bakPath" -ForegroundColor Yellow
        Move-Item-Logged -LiteralPath $offlineWaRoot -Destination $bakPath -Force
    }
    New-Item-Logged -Path $offlineWaRoot -ItemType Directory -Force

    # ------------------------------------------------------------------
    # Step 3: Copy GuestAgent folder(s) to the offline disk
    # ------------------------------------------------------------------
    foreach ($folder in $agentFolders) {
        $dest = Join-Path $offlineWaRoot $folder.Name
        Write-Host "  Copying $($folder.Name) -> $dest" -ForegroundColor Cyan
        Copy-Item-Logged -Path $folder.FullName -Destination $dest -Recurse -Force
    }

    # ------------------------------------------------------------------
    # Step 4 + 5: Import service registry keys into the offline SYSTEM hive
    # ------------------------------------------------------------------
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter 'Windows'
    MountOffHive -WinPath $OfflineWindowsPath -Hive 'SYSTEM'
    try {
        $currentSet = (Get-ItemProperty 'HKLM:\BROKENSYSTEM\Select' -ErrorAction SilentlyContinue).Current
        $csName     = if ($currentSet) { 'ControlSet{0:d3}' -f $currentSet } else { 'ControlSet001' }

        foreach ($svc in $requiredSvcs) {
            $srcReg  = "HKLM\SYSTEM\ControlSet001\Services\$svc"
            $dstReg  = "HKLM\BROKENSYSTEM\$csName\Services\$svc"

            Write-Host "  Copying registry: $svc -> offline $csName..." -ForegroundColor Cyan
            $out = reg.exe copy $srcReg $dstReg /s /f 2>&1
            if ($LASTEXITCODE -ne 0) {
                throw "reg.exe copy failed for '$svc': $out"
            }

            # If the active control set is not ControlSet001, also stamp ControlSet001
            # in the offline hive so both are consistent (Windows reconciles them on boot).
            if ($csName -ne 'ControlSet001') {
                $dstReg001 = "HKLM\BROKENSYSTEM\ControlSet001\Services\$svc"
                if (Test-Path "HKLM:\BROKENSYSTEM\ControlSet001") {
                    $out = reg.exe copy $srcReg $dstReg001 /s /f 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        Write-Warning "  Could not stamp ControlSet001 for '$svc' (may not exist): $out"
                    }
                }
            }

            # Ensure Start = 2 (Automatic) — the copied key may have a different value
            Set-ItemProperty-Logged -Path "HKLM:\BROKENSYSTEM\$csName\Services\$svc" `
                -Name Start -Value 2 -Type DWord -Force
        }

        # RdAgent depends on TermService; make sure it isn't hard-disabled
        $termSvcPath = "HKLM:\BROKENSYSTEM\$csName\Services\TermService"
        if (Test-Path $termSvcPath) {
            $tsBefore = (Get-ItemProperty $termSvcPath -ErrorAction SilentlyContinue).Start
            if ($tsBefore -eq 4) {
                Write-Host "  TermService was Disabled - re-enabling to Auto." -ForegroundColor Yellow
                Set-ItemProperty-Logged -Path $termSvcPath -Name Start -Value 2 -Type DWord -Force
            }
        }

        Write-Host "`nAzure VM Agent installed offline successfully." -ForegroundColor Green
        Write-Host "  Agent binaries : $offlineWaRoot" -ForegroundColor Green
        Write-Host "  Registry keys  : BROKENSYSTEM\$csName\Services\{WindowsAzureGuestAgent,RdAgent}" -ForegroundColor Green
        Write-Host "  After booting the VM, verify the agent with: Get-Service WindowsAzureGuestAgent" -ForegroundColor Cyan
    }
    catch {
        Write-Error "InstallAzureVMAgentOffline failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive 'SYSTEM'
    }
}

function FixDeviceClassFilters {
    param([switch]$KeepDefaultFilters)
    # Scans UpperFilters and LowerFilters for key device class GUIDs identified by Microsoft
    # as critical for boot and network connectivity on Azure VMs.
    #
    # Classes checked (from MS troubleshooting docs):
    #   {4d36e967} DiskDrive  - extra filters  -> stop error 0x7B
    #   {4d36e96a} SCSI/RAID  - extra filters  -> stop error 0x7B
    #   {4d36e97b} SCSIAdapter - extra filters -> stop error 0x7B
    #   {71a27cdd} Volume     - extra filters  -> boot/access failures
    #   {4d36e972} Net        - extra filters  -> complete network loss
    #
    # Decision logic per filter entry (in order of precedence):
    #   1. Binary is missing from offline disk              -> REMOVE (dangling filter will BSOD/hang on boot)
    #   2. Name is in the known-safe list for that class    -> KEEP
    #   3. (skipped when -KeepDefaultFilters) Same entry exists in the host's live class key -> KEEP
    #   4. (skipped when -KeepDefaultFilters) Binary found on offline disk and CompanyName matches Microsoft or an allowed vendor -> KEEP
    #   5. Anything else                                    -> REMOVE and log
    #
    # Special allowances:
    #   Network class: Mellanox/NVIDIA drivers are expected on Azure VMs (Mellanox VF/MANA).
    #
    # GUIDs:
    #   Disk drive     : {4d36e967-e325-11ce-bfc1-08002be10318}
    #   SCSI/RAID      : {4d36e96a-e325-11ce-bfc1-08002be10318}
    #   SCSI Controller: {4d36e97b-e325-11ce-bfc1-08002be10318}
    #   Volume         : {71a27cdd-812a-11d0-bec7-08002be2092f}
    #   Net adapter    : {4d36e972-e325-11ce-bfc1-08002be10318}

    Write-Host "Scanning device class UpperFilters/LowerFilters for unsafe entries..." -ForegroundColor Yellow
    if ($KeepDefaultFilters) {
        Write-Host "  Mode: strict (-KeepDefaultFilters) - only inbox safe-list entries will be kept; Microsoft-signed non-defaults will be removed." -ForegroundColor Cyan
    }
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        $currentSet = (Get-ItemProperty "HKLM:\BROKENSYSTEM\Select" -ErrorAction SilentlyContinue).Current
        $csName     = if ($currentSet) { "ControlSet{0:d3}" -f $currentSet } else { "ControlSet001" }
        $classRoot  = "HKLM:\BROKENSYSTEM\$csName\Control\Class"

        # Class definitions.
        # SafeFilters  : Known Windows inbox / Azure-expected filter names (case-insensitive). Always kept.
        # AllowVendors : Additional vendor substrings matched against binary CompanyName (beyond Microsoft).
        $classChecks = @(
            [PSCustomObject]@{
                GUID         = '{4d36e967-e325-11ce-bfc1-08002be10318}'
                Name         = 'DiskDrive'
                Risk         = 'CRITICAL'
                Description  = 'Extra filters in this class cause stop error 0x7B (INACCESSIBLE_BOOT_DEVICE) on boot'
                # PartMgr     - partition manager (required)
                # fvevol      - BitLocker FVE
                # iorate      - Storage I/O rate limiting (Azure)
                # storqosflt  - Storage QoS filter (Azure)
                # wcifs       - Windows Container Isolation
                # ehstorclass - Enhanced Storage class filter (Microsoft)
                SafeFilters  = [string[]]@('partmgr','fvevol','iorate','storqosflt','wcifs','ehstorclass')
                AllowVendors = [string[]]@()
            },
            [PSCustomObject]@{
                GUID         = '{4d36e96a-e325-11ce-bfc1-08002be10318}'
                Name         = 'SCSIAdapter'
                Risk         = 'CRITICAL'
                Description  = 'Extra filters on SCSI/RAID adapters cause stop error 0x7B (INACCESSIBLE_BOOT_DEVICE)'
                # iasf / iastorf - Intel RST RAID filter (expected on some guest SKUs)
                SafeFilters  = [string[]]@('iasf','iastorf')
                AllowVendors = [string[]]@()
            },
            [PSCustomObject]@{
                GUID         = '{4d36e97b-e325-11ce-bfc1-08002be10318}'
                Name         = 'SCSIController'
                Risk         = 'CRITICAL'
                Description  = 'Extra filters on SCSI controller adapters cause stop error 0x7B (INACCESSIBLE_BOOT_DEVICE)'
                SafeFilters  = [string[]]@()
                AllowVendors = [string[]]@()
            },
            [PSCustomObject]@{
                GUID         = '{71a27cdd-812a-11d0-bec7-08002be2092f}'
                Name         = 'Volume'
                Risk         = 'HIGH'
                Description  = 'Extra filters in this class can cause boot failures or prevent volume access'
                # volsnap    - Volume Shadow Copy (required)
                # fvevol     - BitLocker FVE
                # rdyboost   - ReadyBoost (server: harmless)
                # spldr      - Security processor loader
                # volmgrx    - Volume manager extension
                # iorate / storqosflt - Azure storage filters
                SafeFilters  = [string[]]@('volsnap','fvevol','rdyboost','spldr','volmgrx','iorate','storqosflt')
                AllowVendors = [string[]]@()
            },
            [PSCustomObject]@{
                GUID         = '{4d36e972-e325-11ce-bfc1-08002be10318}'
                Name         = 'Net'
                Risk         = 'HIGH'
                Description  = 'Extra filters in this class can break all network connectivity'
                # WfpLwf                 - Windows Filtering Platform lightweight filter
                # NdisCap                - NDIS capture (Microsoft Network Monitor)
                # NdisImPlatformMpFilter - NDIS IM platform MP filter
                # VmsProxyHNICFilter     - Hyper-V virtual switch proxy
                # vms3cap                - Hyper-V S3 capture filter
                # mslldp                 - Microsoft LLDP Protocol Driver
                # psched                 - QoS Packet Scheduler
                # bridge                 - Network Bridge
                SafeFilters  = [string[]]@('wfplwf','ndiscap','ndisimplatformmpfilter','vmsproxyhnicfilter','vms3cap','mslldp','psched','bridge')
                # Mellanox: Azure VMs use Mellanox ConnectX VF / MANA network adapters.
                # NVIDIA acquired Mellanox; some binaries show either company name.
                AllowVendors = [string[]]@('Mellanox','NVIDIA')
            }
        )

        $anyChanges = $false

        foreach ($classDef in $classChecks) {
            $classPath = "$classRoot\$($classDef.GUID)"
            if (-not (Test-Path $classPath)) {
                Write-Host "  [$($classDef.Name) $($classDef.GUID)] Class key not found in offline hive - skipping." -ForegroundColor DarkGray
                continue
            }

            $riskColor = if ($classDef.Risk -eq 'CRITICAL') { 'Red' } else { 'Yellow' }
            Write-Host "`n  [$($classDef.Name)] $($classDef.Description)" -ForegroundColor Cyan

            # Build host reference path (compare against healthy running system)
            $hostClassPath = "HKLM:\SYSTEM\$csName\Control\Class\$($classDef.GUID)"
            if (-not (Test-Path $hostClassPath)) {
                $hostClassPath = "HKLM:\SYSTEM\ControlSet001\Control\Class\$($classDef.GUID)"
            }

            foreach ($filterType in @('UpperFilters', 'LowerFilters')) {
                $raw = (Get-ItemProperty $classPath -ErrorAction SilentlyContinue).$filterType
                # Normalise: remove nulls/blanks, trim whitespace
                $currentFilters = @($raw | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() })

                if ($currentFilters.Count -eq 0) {
                    Write-Host "    $filterType : (empty)" -ForegroundColor DarkGray
                    continue
                }

                $hostFilters = @()
                if (Test-Path $hostClassPath) {
                    $hostRaw = (Get-ItemProperty $hostClassPath -ErrorAction SilentlyContinue).$filterType
                    $hostFilters = @($hostRaw | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() })
                }

                Write-Host "    $filterType (offline) : $($currentFilters -join ', ')" -ForegroundColor White
                if ($hostFilters.Count -gt 0) {
                    Write-Host "    $filterType (host ref): $($hostFilters -join ', ')" -ForegroundColor DarkGray
                }

                $toKeep  = [System.Collections.Generic.List[string]]::new()
                $removed = [System.Collections.Generic.List[string]]::new()

                foreach ($filter in $currentFilters) {
                    # Resolve the binary path for all checks
                    $svcRegPath  = "HKLM:\BROKENSYSTEM\$csName\Services\$filter"
                    $imgPathRaw  = (Get-ItemProperty $svcRegPath -ErrorAction SilentlyContinue).ImagePath
                    $imgResolved = $null
                    $company     = $null
                    if ($imgPathRaw) {
                        $imgResolved = $imgPathRaw `
                            -replace '(?i)\\SystemRoot\\',  "$script:WinDriveLetter\Windows\" `
                            -replace '(?i)%SystemRoot%',    "$script:WinDriveLetter\Windows" `
                            -replace '(?i)\\\?\?\\',       '' `
                            -replace '(?i)^system32\\',    "$script:WinDriveLetter\Windows\System32\\"
                        if ($imgResolved -match '^(.+?\.(?:sys|exe))') { $imgResolved = $Matches[1] }
                        if (Test-Path $imgResolved) {
                            $company = (Get-Item $imgResolved -ErrorAction SilentlyContinue).VersionInfo.CompanyName
                        }
                    }

                    # Priority 1: binary missing from offline disk -> dangling filter, will BSOD/hang on boot
                    $binaryMissing = ($null -eq $imgPathRaw) -or ($null -ne $imgResolved -and -not (Test-Path $imgResolved))
                    if ($binaryMissing) {
                        $reason = if ($null -eq $imgPathRaw) { 'no ImagePath in registry' } else { "binary not found: $imgResolved" }
                        Write-Host "      [REMOVE] $filter - DANGLING FILTER ($reason) - will cause BSOD/hang on boot" -ForegroundColor Red
                        Write-ActionLog -Event 'DeviceFilterRemoved' -Details @{
                            Class      = $classDef.Name
                            GUID       = $classDef.GUID
                            FilterType = $filterType
                            Filter     = $filter
                            Company    = $null
                            Risk       = $classDef.Risk
                            Reason     = 'DanglingFilter'
                        }
                        $removed.Add($filter)
                        $anyChanges = $true
                        continue
                    }

                    # Priority 2: known-safe list (case-insensitive)
                    if ($classDef.SafeFilters -icontains $filter) {
                        $toKeep.Add($filter)
                        Write-Host "      [KEEP  ] $filter - known safe default" -ForegroundColor Green
                        continue
                    }

                    # Priority 3: present on healthy host reference (skipped in strict mode)
                    if (-not $KeepDefaultFilters -and $hostFilters -icontains $filter) {
                        $toKeep.Add($filter)
                        Write-Host "      [KEEP  ] $filter - present on host reference system" -ForegroundColor Green
                        continue
                    }

                    # Priority 4: binary company name matches an allowed vendor (skipped in strict mode)
                    $isMicrosoftBinary = $company -match 'Microsoft'
                    $isAllowedVendor   = $classDef.AllowVendors | Where-Object { $_ -and $company -match $_ }

                    if (-not $KeepDefaultFilters -and ($isMicrosoftBinary -or $isAllowedVendor)) {
                        $toKeep.Add($filter)
                        Write-Host "      [KEEP  ] $filter - company: '$company'" -ForegroundColor Green
                    } else {
                        # Remove: either strict mode (non-safe-list) or non-Microsoft third-party
                        $vendorInfo = if ($company) { "company: '$company'" } else { 'no version info' }
                        $removeReason = if ($KeepDefaultFilters -and ($isMicrosoftBinary -or $isAllowedVendor)) {
                            "not in safe-list (strict mode)"
                        } else {
                            "non-Microsoft third-party filter ($vendorInfo)"
                        }
                        Write-Host "      [REMOVE] $filter - $removeReason" -ForegroundColor $riskColor
                        Write-ActionLog -Event 'DeviceFilterRemoved' -Details @{
                            Class      = $classDef.Name
                            GUID       = $classDef.GUID
                            FilterType = $filterType
                            Filter     = $filter
                            Company    = $company
                            Risk       = $classDef.Risk
                            Reason     = if ($KeepDefaultFilters) { 'StrictMode-NotInSafeList' } else { 'NonMicrosoftVendor' }
                        }
                        $removed.Add($filter)
                        $anyChanges = $true
                    }
                }

                if ($removed.Count -gt 0) {
                    if ($toKeep.Count -gt 0) {
                        Set-ItemProperty-Logged -Path $classPath -Name $filterType `
                            -Value ([string[]]$toKeep) -Type MultiString -Force
                    } else {
                        # No entries remain - remove the value entirely rather than leaving an empty MultiString
                        Remove-ItemProperty-Logged -Path $classPath -Name $filterType -Force
                    }
                    Write-Host "      Removed: $($removed -join ', ')  |  Kept: $(if ($toKeep.Count) { $toKeep -join ', ' } else { '(none)' })" -ForegroundColor Yellow
                }
            }
        }

        if ($anyChanges) {
            Write-Host "`nDevice class filter cleanup complete." -ForegroundColor Green
            Write-Warning ("If a legitimate product (backup agent, endpoint security, storage appliance) " +
                           "requires a removed filter, restore it from the action log and re-evaluate.")
        } else {
            Write-Host "`nNo unsafe device class filters found. No changes made." -ForegroundColor Green
        }
    }
    catch {
        Write-Error "FixDeviceClassFilters failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

function ScanNetAdapterBindings {
    # Read-only diagnostic: enumerates all installed network binding components
    # (protocols, services, clients) in the offline disk and flags third-party ones.
    # A component is considered third-party when its ComponentId does not start with "ms_".
    # This mirrors the live-system command:
    #   Get-NetAdapterBinding -AllBindings -IncludeHidden | Where ComponentID -notmatch '^ms_'

    Write-Host "Scanning offline disk for third-party network binding components..." -ForegroundColor Yellow
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        $currentSet = (Get-ItemProperty "HKLM:\BROKENSYSTEM\Select" -ErrorAction SilentlyContinue).Current
        $csName     = if ($currentSet) { "ControlSet{0:d3}" -f $currentSet } else { "ControlSet001" }

        # ── Adapter inventory ────────────────────────────────────────────────────
        # Friendly-name lookup: {GUID} -> adapter name (e.g. "Ethernet")
        $adapterNames = @{}
        $netNetKey = "HKLM:\BROKENSYSTEM\$csName\Control\Network\{4D36E972-E325-11CE-BFC1-08002BE10318}"
        if (Test-Path $netNetKey) {
            Get-ChildItem $netNetKey -ErrorAction SilentlyContinue | ForEach-Object {
                $connName = (Get-ItemProperty "$($_.PSPath)\Connection" -ErrorAction SilentlyContinue).Name
                if ($connName) { $adapterNames[$_.PSChildName.ToUpper()] = $connName }
            }
        }

        # Hardware-description lookup: {GUID} -> DriverDesc (e.g. "Mellanox VF")
        $adapterDescs = @{}
        $netClassKey = "HKLM:\BROKENSYSTEM\$csName\Control\Class\{4D36E972-E325-11CE-BFC1-08002BE10318}"
        if (Test-Path $netClassKey) {
            Get-ChildItem $netClassKey -ErrorAction SilentlyContinue |
                Where-Object { $_.PSChildName -match '^\d{4}$' } |
                ForEach-Object {
                    $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
                    if ($p.NetCfgInstanceId) {
                        $adapterDescs[$p.NetCfgInstanceId.ToUpper()] = $p.DriverDesc
                    }
                }
        }

        # Print adapter table
        Write-Host "`nNetwork adapters found in offline disk:" -ForegroundColor Cyan
        $adapterRows = foreach ($guid in ($adapterDescs.Keys | Sort-Object)) {
            $friendlyGuid = $guid   # already upper, with braces or without - normalise
            $cleanGuid = $friendlyGuid.Trim('{','}')
            [PSCustomObject]@{
                FriendlyName = $adapterNames[$cleanGuid]
                Description  = $adapterDescs[$guid]
                GUID         = "{$cleanGuid}"
            }
        }
        if ($adapterRows) {
            $adapterRows | Format-Table FriendlyName, Description, GUID -AutoSize | Out-String | Write-Host
        } else {
            Write-Host "  (none found)" -ForegroundColor DarkGray
        }

        # ── Component enumeration ─────────────────────────────────────────────────
        # Network protocol/service/client software components each live under a
        # numbered instance key inside their Class GUID.
        $componentClasses = @(
            [PSCustomObject]@{ GUID = '{4D36E973-E325-11CE-BFC1-08002BE10318}'; Type = 'Client'   }
            [PSCustomObject]@{ GUID = '{4D36E974-E325-11CE-BFC1-08002BE10318}'; Type = 'Service'  }
            [PSCustomObject]@{ GUID = '{4D36E975-E325-11CE-BFC1-08002BE10318}'; Type = 'Protocol' }
        )

        $allComponents = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($classInfo in $componentClasses) {
            $classKey = "HKLM:\BROKENSYSTEM\$csName\Control\Class\$($classInfo.GUID)"
            if (-not (Test-Path $classKey)) { continue }

            Get-ChildItem $classKey -ErrorAction SilentlyContinue |
                Where-Object { $_.PSChildName -match '^\d{4}$' } |
                ForEach-Object {
                    $props = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
                    # ComponentId casing varies across Windows versions
                    $componentId = $props.ComponentId
                    if (-not $componentId) { $componentId = $props.ComponentID }
                    if (-not $componentId) { return }

                    # Service name: prefer explicit Ndi\Service, fall back to stripping 'ms_' prefix
                    $ndiProps    = Get-ItemProperty "$($_.PSPath)\Ndi" -ErrorAction SilentlyContinue
                    $serviceName = if ($ndiProps -and $ndiProps.Service) { $ndiProps.Service }
                                   else { $componentId -replace '^ms_', '' }

                    # Bound-adapter list from Services\<svc>\Linkage\Bind
                    $boundAdapters   = @()
                    $linkagePath     = "HKLM:\BROKENSYSTEM\$csName\Services\$serviceName\Linkage"
                    if (Test-Path $linkagePath) {
                        $linkage      = Get-ItemProperty $linkagePath -ErrorAction SilentlyContinue
                        $disabledList = @($linkage.Disabled | Where-Object { $_ })
                        foreach ($b in @($linkage.Bind | Where-Object { $_ })) {
                            # Bind entries look like: \Device\{GUID} or \Device\{GUID}_N
                            if ($b -match '\{([0-9A-Fa-f\-]+)\}') {
                                $g    = $Matches[1].ToUpper()
                                $name = $adapterNames[$g]
                                $desc = $adapterDescs["{$g}"]
                                $label = if ($name) { $name } elseif ($desc) { $desc } else { "{$g}" }
                                $boundAdapters += if ($disabledList -icontains $b) { "$label [disabled]" } else { $label }
                            }
                        }
                    }

                    # Binary presence on the offline disk
                    $imgPathRaw  = (Get-ItemProperty "HKLM:\BROKENSYSTEM\$csName\Services\$serviceName" -ErrorAction SilentlyContinue).ImagePath
                    $binaryFound = $null
                    if ($imgPathRaw) {
                        $imgResolved = $imgPathRaw `
                            -replace '(?i)\\SystemRoot\\', "$script:WinDriveLetter\Windows\" `
                            -replace '(?i)%SystemRoot%',   "$script:WinDriveLetter\Windows" `
                            -replace '(?i)\\\?\?\\',      '' `
                            -replace '(?i)^system32\\',   "$script:WinDriveLetter\Windows\System32\\"
                        if ($imgResolved -match '^(.+?\.(?:sys|exe|dll))') { $imgResolved = $Matches[1] }
                        $binaryFound = Test-Path $imgResolved
                    }

                    $allComponents.Add([PSCustomObject]@{
                        ComponentId   = $componentId
                        Type          = $classInfo.Type
                        DisplayName   = $props.DriverDesc
                        ServiceName   = $serviceName
                        IsThirdParty  = ($componentId -notmatch '^ms_')
                        BoundAdapters = ($boundAdapters -join '; ')
                        BinaryFound   = $binaryFound
                    })
                }
        }

        # Deduplicate in case the same ComponentId appears in multiple class GUIDs
        $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $components = $allComponents | Where-Object { $seen.Add($_.ComponentId) }

        $firstParty = @($components | Where-Object { -not $_.IsThirdParty })
        $thirdParty = @($components | Where-Object { $_.IsThirdParty } | Sort-Object ComponentId)

        # ── Report ────────────────────────────────────────────────────────────────
        Write-Host "First-party binding components (ms_*): $($firstParty.Count)" -ForegroundColor DarkGray

        if ($thirdParty.Count -eq 0) {
            Write-Host "`nNo third-party network binding components found." -ForegroundColor Green
        } else {
            Write-Host "`nThird-party network binding components: $($thirdParty.Count)" -ForegroundColor Yellow
            Write-Host "  ComponentId does not start with 'ms_' - these extend the network stack at boot." -ForegroundColor DarkYellow
            Write-Host ""

            foreach ($c in $thirdParty) {
                if ($c.BinaryFound -eq $false) {
                    $presenceTag = 'BINARY MISSING'
                    $color       = 'Red'
                } elseif ($c.BinaryFound -eq $true) {
                    $presenceTag = 'binary present'
                    $color       = 'Yellow'
                } else {
                    $presenceTag = 'no ImagePath registered'
                    $color       = 'DarkYellow'
                }

                Write-Host "  [$($c.Type.PadRight(8))] $($c.ComponentId.PadRight(32)) $($c.DisplayName)" -ForegroundColor $color
                Write-Host "             Service : $($c.ServiceName)  |  $presenceTag" -ForegroundColor $color
                if ($c.BoundAdapters) {
                    Write-Host "             Bound to: $($c.BoundAdapters)" -ForegroundColor DarkYellow
                } else {
                    Write-Host "             Bound to: (linkage not cached offline - resolved at boot)" -ForegroundColor DarkGray
                }
                Write-Host ""
            }

            Write-Warning ("Third-party network components extend the network stack. " +
                           "If the VM has no network connectivity after boot, these components " +
                           "may be interfering. Use -FixDeviceFilters to scan NDIS UpperFilters/LowerFilters.")
        }

        Write-ActionLog -Event 'ScanNetAdapterBindings' -Details @{
            AdapterCount    = $adapterRows.Count
            FirstPartyCount = $firstParty.Count
            ThirdPartyCount = $thirdParty.Count
            ThirdPartyIds   = ($thirdParty | ForEach-Object ComponentId) -join ', '
        }
    }
    catch {
        Write-Error "ScanNetAdapterBindings failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

function RemoveOrphanedNetBindings {
    # Removes third-party network binding components whose driver binary is missing on the offline disk.
    # A dangling binding component causes ndis.sys initialisation to fail on boot, leading to complete
    # loss of network connectivity even when the NIC driver itself is healthy.
    #
    # This is the destructive counterpart to -ScanNetBindings. Run that first to preview.
    # Only components with an explicit ImagePath that resolves to a non-existent file are acted on;
    # components with no ImagePath at all are left untouched (they may be imageless by design).
    #
    # For each orphaned component:
    #   1. Removes the class instance key (Control\Class\{type-GUID}\NNNN)
    #   2. Disables the service (Services\<svc> Start = 4) to prevent any load attempt
    #   3. Clears Services\<svc>\Linkage\{Bind, Export, Route} to detach from all adapters

    Write-Host "Scanning for orphaned third-party network binding components (missing binary)..." -ForegroundColor Yellow
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        $currentSet = (Get-ItemProperty "HKLM:\BROKENSYSTEM\Select" -ErrorAction SilentlyContinue).Current
        $csName     = if ($currentSet) { "ControlSet{0:d3}" -f $currentSet } else { "ControlSet001" }

        $componentClasses = @(
            [PSCustomObject]@{ GUID = '{4D36E973-E325-11CE-BFC1-08002BE10318}'; Type = 'Client'   }
            [PSCustomObject]@{ GUID = '{4D36E974-E325-11CE-BFC1-08002BE10318}'; Type = 'Service'  }
            [PSCustomObject]@{ GUID = '{4D36E975-E325-11CE-BFC1-08002BE10318}'; Type = 'Protocol' }
        )

        $orphans = [System.Collections.Generic.List[PSCustomObject]]::new()
        $seen    = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

        foreach ($classInfo in $componentClasses) {
            $classKey = "HKLM:\BROKENSYSTEM\$csName\Control\Class\$($classInfo.GUID)"
            if (-not (Test-Path $classKey)) { continue }

            Get-ChildItem $classKey -ErrorAction SilentlyContinue |
                Where-Object { $_.PSChildName -match '^\d{4}$' } |
                ForEach-Object {
                    $instPath    = $_.PSPath
                    $props       = Get-ItemProperty $instPath -ErrorAction SilentlyContinue
                    $componentId = if ($props.ComponentId) { $props.ComponentId } else { $props.ComponentID }
                    if (-not $componentId) { return }
                    # Deduplicate across class GUIDs
                    if (-not $seen.Add($componentId)) { return }
                    # Never touch first-party ms_ components
                    if ($componentId -match '^ms_') { return }

                    $ndiProps    = Get-ItemProperty "$instPath\Ndi" -ErrorAction SilentlyContinue
                    $serviceName = if ($ndiProps -and $ndiProps.Service) { $ndiProps.Service }
                                   else { $componentId -replace '^ms_', '' }

                    $imgPathRaw = (Get-ItemProperty "HKLM:\BROKENSYSTEM\$csName\Services\$serviceName" -ErrorAction SilentlyContinue).ImagePath
                    # No ImagePath -> skip; component may be imageless by design
                    if ($null -eq $imgPathRaw) { return }

                    $imgResolved = $imgPathRaw `
                        -replace '(?i)\\SystemRoot\\', "$script:WinDriveLetter\Windows\" `
                        -replace '(?i)%SystemRoot%',   "$script:WinDriveLetter\Windows" `
                        -replace '(?i)\\\?\?\\',      '' `
                        -replace '(?i)^system32\\',   "$script:WinDriveLetter\Windows\System32\\"
                    if ($imgResolved -match '^(.+?\.(?:sys|exe|dll))') { $imgResolved = $Matches[1] }
                    # Binary present -> nothing to do
                    if (Test-Path $imgResolved) { return }

                    $orphans.Add([PSCustomObject]@{
                        ComponentId   = $componentId
                        Type          = $classInfo.Type
                        DisplayName   = $props.DriverDesc
                        ServiceName   = $serviceName
                        ClassKeyPath  = $instPath
                        MissingBinary = $imgResolved
                    })
                }
        }

        if ($orphans.Count -eq 0) {
            Write-Host "`nNo orphaned network binding components found. No changes made." -ForegroundColor Green
            return
        }

        Write-Host "`nFound $($orphans.Count) orphaned network binding component(s) with missing binaries:" -ForegroundColor Red
        Write-Host ""

        foreach ($c in $orphans) {
            Write-Host "  [$($c.Type.PadRight(8))] $($c.ComponentId)  --  $($c.DisplayName)" -ForegroundColor Red
            Write-Host "             Service : $($c.ServiceName)" -ForegroundColor DarkRed
            Write-Host "             Missing : $($c.MissingBinary)" -ForegroundColor DarkRed
            Write-Host ""

            # Step 1: Remove class instance key - de-registers the component from Windows
            Write-Host "    [1/3] Removing class instance key..." -ForegroundColor Yellow
            Remove-Item-Logged -Path $c.ClassKeyPath -Recurse -Force

            # Step 2: Disable the service to prevent any load attempt at boot
            $svcPath = "HKLM:\BROKENSYSTEM\$csName\Services\$($c.ServiceName)"
            if (Test-Path $svcPath) {
                Write-Host "    [2/3] Disabling service '$($c.ServiceName)' (Start -> 4)..." -ForegroundColor Yellow
                Set-ItemProperty-Logged -Path $svcPath -Name Start -Value 4 -Type DWord -Force
            } else {
                Write-Host "    [2/3] Service key '$($c.ServiceName)' not found - skipping." -ForegroundColor DarkGray
            }

            # Step 3: Clear Linkage Bind/Export/Route to detach from all adapters
            $linkPath = "HKLM:\BROKENSYSTEM\$csName\Services\$($c.ServiceName)\Linkage"
            if (Test-Path $linkPath) {
                Write-Host "    [3/3] Clearing Linkage bind entries..." -ForegroundColor Yellow
                foreach ($val in @('Bind', 'Export', 'Route')) {
                    if ($null -ne (Get-ItemProperty $linkPath -ErrorAction SilentlyContinue).$val) {
                        Remove-ItemProperty-Logged -Path $linkPath -Name $val -Force
                    }
                }
            } else {
                Write-Host "    [3/3] No Linkage key found - nothing to clear." -ForegroundColor DarkGray
            }

            Write-ActionLog -Event 'OrphanedNetBindingRemoved' -Details @{
                ComponentId   = $c.ComponentId
                Type          = $c.Type
                DisplayName   = $c.DisplayName
                ServiceName   = $c.ServiceName
                MissingBinary = $c.MissingBinary
            }

            Write-Host "    Done: $($c.ComponentId) removed." -ForegroundColor Green
            Write-Host ""
        }

        Write-Host "Orphaned network binding cleanup complete. $($orphans.Count) component(s) removed." -ForegroundColor Green
        Write-Warning ("A full network-stack reset may still be needed. " +
                       "Consider also running -ResetNetworkStack to clear TCP/IP and Winsock state.")
    }
    catch {
        Write-Error "RemoveOrphanedNetBindings failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

function CopyACPISettings {
    # Copies legacy Hyper-V ACPI Enum entries to the newer ACPI device IDs.
    # Some Windows versions do not include the newer MSFT* keys, preventing
    # Hyper-V synthetic devices from being detected. This copies from the
    # legacy entries when they exist:
    #   VMBus                  -> MSFT1000
    #   Hyper_V_Gen_Counter_V1 -> MSFT1002

    Write-Host "Copying Hyper-V ACPI device entries to newer device IDs..." -ForegroundColor Cyan

    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        $sysRoot = Get-SystemRootPath   # e.g. HKLM:\BROKENSYSTEM\ControlSet001
        # Convert PowerShell path to reg.exe syntax (no colon after HKLM)
        $regRoot = $sysRoot -replace '^HKLM:', 'HKLM'

        $copies = @(
            @{ SrcPs = "$sysRoot\Enum\ACPI\VMBus";                  DstPs = "$sysRoot\Enum\ACPI\MSFT1000"; SrcReg = "$regRoot\Enum\ACPI\VMBus";                  DstReg = "$regRoot\Enum\ACPI\MSFT1000"; Label = 'VMBus -> MSFT1000' }
            @{ SrcPs = "$sysRoot\Enum\ACPI\Hyper_V_Gen_Counter_V1"; DstPs = "$sysRoot\Enum\ACPI\MSFT1002"; SrcReg = "$regRoot\Enum\ACPI\Hyper_V_Gen_Counter_V1"; DstReg = "$regRoot\Enum\ACPI\MSFT1002"; Label = 'Hyper_V_Gen_Counter_V1 -> MSFT1002' }
        )

        $copied = 0
        foreach ($c in $copies) {
            if (-not (Test-Path $c.SrcPs)) {
                Write-Warning "  Source key not found: $($c.SrcReg) - skipping $($c.Label)"
                continue
            }
            # Run as SYSTEM to bypass ACL restrictions on the mounted offline hive
            ExecuteAsSystem "reg.exe copy `"$($c.SrcReg)`" `"$($c.DstReg)`" /s /f"
            # Verify the destination key was created
            if (Test-Path $c.DstPs) {
                Write-Host "  [OK] Copied $($c.Label)" -ForegroundColor Green
                $copied++
            } else {
                Write-Warning "  reg copy $($c.Label) may have failed - destination key not found after copy"
            }
        }

        if ($copied -gt 0) {
            Write-Host "ACPI settings copied ($copied of $($copies.Count) pair(s))." -ForegroundColor Green
        } else {
            Write-Warning "No ACPI entries were copied - source keys may not exist on this disk."
        }

        Write-ActionLog -Event 'CopyACPISettings' -Details @{ RegRoot = $regRoot; Copied = $copied }
    }
    catch {
        Write-Error "CopyACPISettings failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

function RunSystemCheck {
    # ─────────────────────────────────────────────────────────────────────────
    # Read-only offline health scan. Checks BCD, registry, services, device
    # filters, networking, RDP, Azure Agent, security settings and crash
    # artifacts. No changes are made. At the end a prioritised summary is
    # printed with the exact -Parameter to run for each finding.
    # ─────────────────────────────────────────────────────────────────────────

    # Resolve guest computer name from the offline SYSTEM hive
    $guestComputerName = ''
    try {
        $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
        MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
        try {
            $sysRoot = Get-SystemRootPath
            $guestComputerName = (Get-ItemProperty "$sysRoot\Control\ComputerName\ComputerName" -ErrorAction SilentlyContinue).ComputerName
        } finally {
            UnmountOffHive -Hive "SYSTEM"
        }
    } catch { $guestComputerName = '' }
    $guestLabel = if ($guestComputerName) { "  Guest: $guestComputerName" } else { '' }

    Write-Host "`n===================================================================" -ForegroundColor Cyan
    Write-Host "  Offline System Health Check" -ForegroundColor Cyan
    Write-Host "  Disk $script:DiskNumber  |  Windows: $script:WinDriveLetter  |  Boot: $script:BootDriveLetter  |  Gen$script:VMGen$guestLabel" -ForegroundColor Cyan
    Write-Host "===================================================================`n" -ForegroundColor Cyan

    $findings = [System.Collections.Generic.List[PSCustomObject]]::new()
    $scriptFile = if ($PSCommandPath) { Split-Path -Leaf $PSCommandPath } else { 'Repair-AzVMDisk.ps1' }

    # Inline helper: emit one finding to the console and add it to $findings.
    # Called as: & $emit 'Category' 'CRIT|WARN|INFO|OK' 'Message' '-FixParam'
    $emit = {
        param([string]$Cat, [string]$Sev, [string]$Desc, [string]$Fix = '')
        $findings.Add([PSCustomObject]@{ Category=$Cat; Severity=$Sev; Description=$Desc; Fix=$Fix })
        $color  = switch ($Sev) { 'CRIT' {'Red'} 'WARN' {'Yellow'} 'INFO' {'Cyan'} default {'Green'} }
        $prefix = switch ($Sev) { 'CRIT' {'[CRIT]'} 'WARN' {'[WARN]'} 'INFO' {'[INFO]'} default {'[ OK ]'} }
        Write-Host "  $prefix [$Cat] $Desc" -ForegroundColor $color
        if ($Fix) { Write-Host "         Suggestion: $Fix" -ForegroundColor DarkCyan }
    }

    # ── SysCheck severity configuration ──────────────────────────────────────
    # Controls how each finding is surfaced when it fires.
    # Values: 0 = INFO  |  1 = WARN  |  2 = CRIT
    # Change a value here to promote or demote any test without touching its logic.

    # Disk & Filesystem
    $sevDiskHealth           = 2   # Disk HealthStatus is not Healthy
    $sevDiskRawFs            = 2   # Partition has RAW filesystem (unreadable)
    $sevDiskFsHealth         = 1   # Partition filesystem is unhealthy

    # Crash & Boot artefacts
    $sevCrashMinidumps       = 1   # Minidump (.dmp) files found
    $sevBootNtbtlog          = 0   # ntbtlog.txt present (check for DIDNOTLOAD)
    $sevUpdatePendingXml     = 1   # pending.xml found (update boot loop risk)
    $sevUpdateTxRLogs        = 1   # TxR transaction log files found
    $sevUpdateSmiLogs        = 1   # SMI Store transaction log files found
    $sevSetupMode            = 0   # SetupType active (Setup CmdLine will run at boot)

    # BCD / Boot Configuration
    $sevBcdMissing           = 2   # BCD store not found
    $sevBcdNoBootLoader      = 2   # No Windows Boot Loader entry in BCD
    $sevBcdSafeMode          = 1   # Safe Mode flag active in BCD
    $sevBcdBootStatusPolicy  = 0   # bootstatuspolicy IgnoreAllFailures set
    $sevBcdRecoveryDisabled  = 0   # recoveryenabled Off
    $sevBcdTestSigning       = 1   # Test signing ON
    $sevBcdUnknownDevice     = 2   # BCD contains unknown device/path entries

    # Registry
    $sevControlSetMismatch   = 1   # Current ControlSet != Default
    $sevRegBackEmpty         = 1   # RegBack\SYSTEM is 0 bytes
    $sevRegBackMissing       = 0   # RegBack\SYSTEM not found

    # Critical Services
    $sevCriticalSvcDisabled  = 2   # A critical boot/system driver is disabled

    # RDP
    $sevRdpDenied            = 2   # fDenyTSConnections=1 (RDP explicitly disabled)
    $sevRdpDenyUnknown       = 1   # fDenyTSConnections key missing
    $sevRdpSvcDisabledCrit   = 2   # TermService/SessionEnv disabled
    $sevRdpSvcDisabledWarn   = 1   # UmRdpService disabled
    $sevRdpNonDefaultPort    = 1   # RDP port is not 3389
    $sevRdpSecurityLayerWeak = 1   # SecurityLayer=0 (RDP native, no SSL)
    $sevRdpNLADisabled       = 1   # NLA/UserAuthentication disabled
    $sevRdpSecProtoNeg       = 1   # fAllowSecProtocolNegotiation=0
    $sevRdpMinEncLevel       = 1   # MinEncryptionLevel below 2
    $sevRdpTcpKeyMissing     = 1   # RDP-Tcp WinStation key not found
    $sevRdpCryptoSvcDisabled = 1   # KeyIso/CryptSvc/CertPropSvc disabled
    $sevRdpTlsDisabled       = 1   # TLS 1.2 explicitly disabled in SCHANNEL
    $sevRdpNtlmRestrict      = 1   # NTLM restrictions may block RDP auth
    $sevRdpLmCompat          = 1   # LmCompatibilityLevel > 5
    $sevCredSspOracle        = 1   # CredSSP AllowEncryptionOracle != 2
    $sevGpRdpBlocked         = 2   # Group Policy is blocking RDP
    $sevGpNlaDisabled        = 0   # Group Policy has disabled NLA
    $sevSslCipherPolicy      = 1   # SSL cipher suite policy configured
    $sevRdpKeySystemAcl      = 1   # RDP private key: SYSTEM missing FullControl
    $sevRdpKeyNetSvcAcl      = 1   # RDP private key: NETWORK SERVICE missing Read
    $sevRdpKeySessionEnvAcl  = 0   # RDP private key: SessionEnv missing FullControl
    $sevRdpKeyFileMissing    = 1   # RDP private key file not found in MachineKeys
    $sevRdpKeyZeroLength     = 1   # Zero-length files in MachineKeys
    $sevMachineKeysMissing   = 1   # MachineKeys folder missing

    # Security
    $sevCredentialGuard      = 1   # Credential Guard is enabled
    $sevLsaPPL               = 0   # LSA RunAsPPL=1 active
    $sevAppLockerEnforcing   = 1   # AppLocker is enforcing rules
    $sevAppIdSvc             = 0   # AppIDSvc (Application Identity) is running

    # Azure Guest Agent
    $sevAzureAgentDisabled   = 2   # Agent service is disabled
    $sevAzureAgentWrongStart = 1   # Agent service has unexpected start type
    $sevAzureAgentMissing    = 1   # Agent not found in registry or on disk

    # Networking
    $sevBfeDisabled          = 1   # BFE (Base Filtering Engine) disabled
    $sevTcpipDisabled        = 2   # Tcpip service disabled
    $sevNetSvcDisabled       = 1   # Secondary networking services disabled (DNS/DHCP/NLA/SMB)
    $sevSanPolicy            = 1   # SAN policy is not OnlineAll
    $sevOrphanedNdis         = 2   # Orphaned NDIS bindings with missing binary

    # Drivers
    $sevMissingDriverBinaries = 2  # Boot/System driver registered but binary missing

    # Device class filters
    $sevDeviceFiltersCrit    = 2   # Non-standard entries in DiskDrive/SCSI class filters
    $sevDeviceFiltersWarn    = 1   # Non-standard entries in Volume/Net class filters

    # Windows Update
    $sevUpdateWuDisabled     = 0   # WU services (wuauserv/UsoSvc/WaaSMedicSvc) disabled
    $sevCbsPendingWarn       = 1   # CBS RebootPending/PackagesPending/exclusive SessionsPending
    $sevCbsPendingInfo       = 0   # CBS SessionsPending without exclusive lock (stale, usually benign)

    # ACPI
    $sevACPISettings         = 0   # Hyper-V ACPI entries (MSFT1000/MSFT1002) missing

    # Inline converter: integer level -> severity string used by $emit
    $toSev = { param([int]$L) switch ($L) { 0{'INFO'} 1{'WARN'} 2{'CRIT'} default{'INFO'} } }

    # ── 1. Disk & Filesystem ─────────────────────────────────────────────────
    Write-Host "--- Disk & Filesystem" -ForegroundColor DarkGray
    try {
        $disk = Get-Disk -Number $script:DiskNumber -ErrorAction SilentlyContinue
        if ($disk) {
            if ($disk.HealthStatus -ne 'Healthy') {
                & $emit 'Disk' (& $toSev $sevDiskHealth) "Disk health: $($disk.HealthStatus)" "-CheckDiskHealth"
            } else {
                & $emit 'Disk' 'OK' "Disk health: Healthy"
            }
        }
        foreach ($p in (Get-Partition -DiskNumber $script:DiskNumber -ErrorAction SilentlyContinue)) {
            $vol = Get-Volume -Partition $p -ErrorAction SilentlyContinue
            if ($vol) {
                $fs = $vol.FileSystemType
                $pLetter = if ($p.DriveLetter -and $p.DriveLetter -ne "`0") { "$($p.DriveLetter):" } else { $script:WinDriveLetter.TrimEnd('\') }
                if ([string]::IsNullOrEmpty($fs) -or $fs -eq 'Unknown') {
                    & $emit 'Disk' (& $toSev $sevDiskRawFs) "Partition $($p.PartitionNumber) ($pLetter): RAW filesystem - data may be inaccessible" "-FixNTFS -DriveLetter $pLetter -LeaveDiskOnline"
                } elseif ($vol.HealthStatus -ne 'Healthy') {
                    & $emit 'Disk' (& $toSev $sevDiskFsHealth) "Partition $($p.PartitionNumber) ($pLetter) ($fs): health=$($vol.HealthStatus)" "-FixNTFS -DriveLetter $pLetter -LeaveDiskOnline"
                }
            }
        }
    } catch { Write-Warning "Disk check failed: $_" }

    # ── 2. Crash & Boot Artefacts ────────────────────────────────────────────
    Write-Host "--- Crash & Boot Artefacts" -ForegroundColor DarkGray
    $minidumpDir = Join-Path $script:WinDriveLetter 'Windows\Minidump'
    if (Test-Path $minidumpDir) {
        $dumps = @(Get-ChildItem $minidumpDir -Filter '*.dmp' -ErrorAction SilentlyContinue)
        if ($dumps.Count -gt 0) {
            $newest = $dumps | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            & $emit 'Crash' (& $toSev $sevCrashMinidumps) "$($dumps.Count) minidump(s) found - latest: $($newest.Name) [$($newest.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))]" "-CollectEventLogs"
        } else { & $emit 'Crash' 'OK' 'No minidump files' }
    } else { & $emit 'Crash' 'OK' 'No Minidump folder' }

    if (Test-Path (Join-Path $script:WinDriveLetter 'Windows\ntbtlog.txt')) {
        & $emit 'Boot' (& $toSev $sevBootNtbtlog) 'ntbtlog.txt present - check for DIDNOTLOAD entries' "-CollectEventLogs"
    }

    if (Test-Path (Join-Path $script:WinDriveLetter 'Windows\WinSxS\pending.xml')) {
        & $emit 'WindowsUpdate' (& $toSev $sevUpdatePendingXml) 'Pending Windows Update transaction (pending.xml) - may cause boot loop on Configuring Updates screen' "-FixPendingUpdates"
    }

    # TxR transaction log files (leftover .blf/.regtrans-ms cause stuck update processing)
    $txrFolder = Join-Path $script:WinDriveLetter 'Windows\System32\config\TxR'
    if (Test-Path $txrFolder) {
        $txrFiles = @(Get-ChildItem $txrFolder -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.blf','.regtrans-ms') })
        if ($txrFiles.Count -gt 0) {
            & $emit 'WindowsUpdate' (& $toSev $sevUpdateTxRLogs) "$($txrFiles.Count) TxR transaction log file(s) found in config\TxR (.blf/.regtrans-ms) - may cause update processing to hang at boot" "-FixPendingUpdates"
        }
    }

    # SMI Store transaction files
    $smiFolder = Join-Path $script:WinDriveLetter 'Windows\System32\SMI\Store\Machine'
    if (Test-Path $smiFolder) {
        $smiFiles = @(Get-ChildItem $smiFolder -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.blf','.regtrans-ms') })
        if ($smiFiles.Count -gt 0) {
            & $emit 'WindowsUpdate' (& $toSev $sevUpdateSmiLogs) "$($smiFiles.Count) SMI Store transaction log file(s) found (.blf/.regtrans-ms) - may cause update boot loop" "-FixPendingUpdates"
        }
    }

    # ── 3. BCD ───────────────────────────────────────────────────────────────
    Write-Host "--- BCD / Boot Configuration" -ForegroundColor DarkGray
    $bcdPath = Get-BcdStorePath -Generation $script:VMGen -BootDrive $script:BootDriveLetter.TrimEnd('\')
    if (-not (Test-Path $bcdPath)) {
        & $emit 'BCD' (& $toSev $sevBcdMissing) "BCD store not found at $bcdPath - VM will fail to boot" "-FixBoot"
    } else {
        & $emit 'BCD' 'OK' "BCD store present: $bcdPath"
        try {
            $bcdText = (& bcdedit.exe /store "$bcdPath" /enum all 2>&1) | Out-String
            if ($bcdText -notmatch 'Windows Boot Loader|osloader') {
                & $emit 'BCD' (& $toSev $sevBcdNoBootLoader) 'No Windows Boot Loader entry found in BCD' "-FixBoot"
            } else {
                & $emit 'BCD' 'OK' 'Windows Boot Loader entry present'
            }
            if ($bcdText -match 'safeboot\s+(\S+)') {
                & $emit 'BCD' (& $toSev $sevBcdSafeMode) "Safe Mode boot flag is active (safeboot $($Matches[1])) - VM will boot into Safe Mode" "-RemoveSafeModeFlag"
            }
            if ($bcdText -match 'bootstatuspolicy\s+ignoreallfailures') {
                & $emit 'BCD' (& $toSev $sevBcdBootStatusPolicy) 'bootstatuspolicy IgnoreAllFailures is set - startup repair is suppressed'
            }
            if ($bcdText -match 'recoveryenabled\s+no') {
                & $emit 'BCD' (& $toSev $sevBcdRecoveryDisabled) 'recoveryenabled is Off - WinRE recovery disabled'
            }
            if ($bcdText -match 'testsigning\s+yes') {
                & $emit 'BCD' (& $toSev $sevBcdTestSigning) 'Test signing is ON - unsigned drivers are permitted to load' "-DisableTestSigning"
            }
            if ($bcdText -match '\bunknown\b') {
                & $emit 'BCD' (& $toSev $sevBcdUnknownDevice) 'BCD contains entries with unknown device/path - may point to wrong or missing partition' "-FixBoot"
            }
        } catch { & $emit 'BCD' 'WARN' "Could not enumerate BCD: $_" }
    }

    # ── 4. SYSTEM hive ───────────────────────────────────────────────────────
    Write-Host "--- Registry / Services (SYSTEM hive)" -ForegroundColor DarkGray
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter 'Windows'
    MountOffHive -WinPath $OfflineWindowsPath -Hive 'SYSTEM'
    try {
        $sel     = Get-ItemProperty 'HKLM:\BROKENSYSTEM\Select' -ErrorAction SilentlyContinue
        $curSet  = $sel.Current
        $lkgcSet = $sel.LastKnownGood
        $defSet  = $sel.Default
        $csName  = if ($curSet) { 'ControlSet{0:d3}' -f $curSet } else { 'ControlSet001' }
        $svcRoot = "HKLM:\BROKENSYSTEM\$csName\Services"
        $ctrlRoot= "HKLM:\BROKENSYSTEM\$csName\Control"

        # ControlSet mismatch
        if ($null -ne $curSet -and $null -ne $defSet -and $curSet -ne $defSet) {
            & $emit 'Registry' (& $toSev $sevControlSetMismatch) "ControlSet mismatch: Current=$curSet Default=$defSet - LKGC switch may help" "-TryLGKC"
        } else {
            & $emit 'Registry' 'OK' "ControlSet: Current=ControlSet$("{0:d3}" -f $curSet)  LKGC=ControlSet$("{0:d3}" -f $lkgcSet)"
        }

        # RegBack
        $rbSystem = Join-Path $script:WinDriveLetter 'Windows\System32\config\RegBack\SYSTEM'
        if (Test-Path $rbSystem) {
            $rbSize = (Get-Item $rbSystem).Length
            if ($rbSize -eq 0) {
                & $emit 'Registry' (& $toSev $sevRegBackEmpty) 'RegBack\SYSTEM is 0 bytes - no registry backup is available' "-EnableRegBackup"
            } else {
                & $emit 'Registry' 'OK' "RegBack\SYSTEM is $([math]::Round($rbSize/1MB, 1)) MB"
            }
        } else {
            & $emit 'Registry' (& $toSev $sevRegBackMissing) 'RegBack\SYSTEM not found - registry backups not configured' "-EnableRegBackup"
        }

        # Pending Setup CmdLine
        $setupProps = Get-ItemProperty 'HKLM:\BROKENSYSTEM\Setup' -ErrorAction SilentlyContinue
        if ($setupProps.SetupType -and $setupProps.SetupType -ne 0) {
            & $emit 'Boot' (& $toSev $sevSetupMode) "Setup mode active (SetupType=$($setupProps.SetupType)) - VM will run: '$($setupProps.CmdLine)' on next boot"
        }

        # ── Critical boot services ──────────────────────────────────────────
        # These disabled = guaranteed BSOD 0x7B or non-boot
        $critical = @(
            @{ N='disk';     ExpStart=0; Desc='storage bus driver (0x7B if disabled)' }
            @{ N='volmgr';   ExpStart=0; Desc='volume manager (0x7B if disabled)' }
            @{ N='partmgr';  ExpStart=1; Desc='partition manager (0x7B if disabled)' }
            @{ N='storport'; ExpStart=0; Desc='storage port driver (0x7B if disabled)' }
            @{ N='NTFS';     ExpStart=1; Desc='NTFS filesystem driver (0x7B if disabled)' }
            @{ N='volsnap';  ExpStart=1; Desc='volume shadow copy filter' }
            @{ N='msrpc';    ExpStart=2; Desc='RPC subsystem' }
            @{ N='rpcss';    ExpStart=2; Desc='RPC Endpoint Mapper' }
            @{ N='LSM';      ExpStart=2; Desc='Local Session Manager' }
        )
        foreach ($s in $critical) {
            $sp = "$svcRoot\$($s.N)"
            if (Test-Path $sp) {
                $start = (Get-ItemProperty $sp -ErrorAction SilentlyContinue).Start
                if ($start -eq 4) {
                    & $emit 'Services' (& $toSev $sevCriticalSvcDisabled) "$($s.N) is DISABLED (Start=4) - $($s.Desc) [ re-enable: Set Start=$($s.ExpStart) ]" "-EnableDriver $($s.N)"
                }
            }
        }

        # ── RDP ────────────────────────────────────────────────────────────
        $tsPath     = "$ctrlRoot\Terminal Server"
        $rdpTcpPath = "$ctrlRoot\Terminal Server\WinStations\RDP-Tcp"
        $fDeny = (Get-ItemProperty $tsPath -ErrorAction SilentlyContinue).fDenyTSConnections
        if ($fDeny -eq 1) {
            & $emit 'RDP' (& $toSev $sevRdpDenied) 'fDenyTSConnections=1 - RDP is disabled at canonical key' "-FixRDP"
        } elseif ($null -eq $fDeny) {
            & $emit 'RDP' (& $toSev $sevRdpDenyUnknown) 'fDenyTSConnections not found - RDP state unclear' "-FixRDP"
        } else {
            & $emit 'RDP' 'OK' 'fDenyTSConnections=0 - RDP is enabled'
        }

        # TermService, SessionEnv, UmRdpService
        foreach ($rdpSvc in @(
            @{ N='TermService';  Desc='Remote Desktop Services'; Crit=$true }
            @{ N='SessionEnv';   Desc='Remote Desktop Config';   Crit=$true }
            @{ N='UmRdpService'; Desc='RDP UserMode Port Redirector'; Crit=$false }
        )) {
            $svcStart = (Get-ItemProperty "$svcRoot\$($rdpSvc.N)" -ErrorAction SilentlyContinue).Start
            if ($svcStart -eq 4) {
                $sev = if ($rdpSvc.Crit) { (& $toSev $sevRdpSvcDisabledCrit) } else { (& $toSev $sevRdpSvcDisabledWarn) }
                & $emit 'RDP' $sev "$($rdpSvc.N) ($($rdpSvc.Desc)) is DISABLED (Start=4) - RDP will not work" "-FixRDP"
            } elseif ($null -ne $svcStart) {
                & $emit 'RDP' 'OK' "$($rdpSvc.N) Start=$svcStart"
            }
        }

        if (Test-Path $rdpTcpPath) {
            $rdpP = Get-ItemProperty $rdpTcpPath -ErrorAction SilentlyContinue

            # Port number
            $rdpPort = $rdpP.PortNumber
            if ($null -ne $rdpPort -and $rdpPort -ne 3389) {
                & $emit 'RDP' (& $toSev $sevRdpNonDefaultPort) "RDP-Tcp port is $rdpPort (not the default 3389) - ensure firewall allows this port or run -FixRDP to reset" "-FixRDP"
            } else {
                & $emit 'RDP' 'OK' "RDP-Tcp port: $(if ($null -eq $rdpPort) { '(not set, default 3389)' } else { $rdpPort })"
            }

            # Security layer
            $sl = $rdpP.SecurityLayer
            if ($null -ne $sl) {
                $slDesc = switch ($sl) { 0 {'RDP native (no SSL) - weakest encryption'} 1 {'Negotiate'} 2 {'SSL/TLS required'} default {"Unknown ($sl)"} }
                $slSev  = if ($sl -eq 0) { (& $toSev $sevRdpSecurityLayerWeak) } else { 'INFO' }
                & $emit 'RDP' $slSev "SecurityLayer=$sl ($slDesc)"
            }

            # NLA / UserAuthentication
            $ua = $rdpP.UserAuthentication
            if ($ua -eq 0) {
                & $emit 'RDP' (& $toSev $sevRdpNLADisabled) 'NLA is DISABLED (UserAuthentication=0) - any user can attempt login without pre-auth; run -EnableNLA to restore' "-EnableNLA"
            } elseif ($ua -eq 1) {
                & $emit 'RDP' 'OK' 'NLA is enabled (UserAuthentication=1)'
            } else {
                & $emit 'RDP' 'INFO' 'NLA/UserAuthentication not set on RDP-Tcp key (policy may control this)'
            }

            # fAllowSecProtocolNegotiation (also touched by -DisableNLA)
            $aspn = $rdpP.fAllowSecProtocolNegotiation
            if ($aspn -eq 0) {
                & $emit 'RDP' (& $toSev $sevRdpSecProtoNeg) 'fAllowSecProtocolNegotiation=0 - security protocol negotiation disabled; run -EnableNLA or -FixRDP to reset' "-EnableNLA"
            }

            # MinEncryptionLevel (1=Low was set by -DisableNLA; default is typically 2)
            $mel = $rdpP.MinEncryptionLevel
            if ($null -ne $mel -and $mel -lt 2) {
                & $emit 'RDP' (& $toSev $sevRdpMinEncLevel) "MinEncryptionLevel=$mel (below recommended 2) - may indicate NLA was disabled" "-EnableNLA"
            }

            # SSL certificate thumbprint - absence is normal (Windows generates one on first RDP connection)
            $cert = $rdpP.SSLCertificateSHA1Hash
            if ($null -ne $cert -and -not ($cert -is [byte[]] -and $cert.Count -eq 0)) {
                & $emit 'RDP' 'OK' 'RDP SSL certificate thumbprint is present'
            }
            # No thumbprint = not flagged; Windows auto-generates one at first RDP connection
        } else {
            & $emit 'RDP' (& $toSev $sevRdpTcpKeyMissing) 'RDP-Tcp WinStation key not found - RDP listener may be misconfigured' "-FixRDP"
        }

        # RDP-related crypto services (required for certificate/key operations)
        foreach ($cryptSvc in @(
            @{ N='KeyIso';      DefStart=3; Desc='CNG Key Isolation (needed for RDP private key)' }
            @{ N='CryptSvc';    DefStart=2; Desc='Cryptographic Services (needed for cert store)' }
            @{ N='CertPropSvc'; DefStart=3; Desc='Certificate Propagation (needed for user certs)' }
        )) {
            $cs = (Get-ItemProperty "$svcRoot\$($cryptSvc.N)" -ErrorAction SilentlyContinue).Start
            if ($cs -eq 4) {
                & $emit 'RDP' (& $toSev $sevRdpCryptoSvcDisabled) "$($cryptSvc.N) ($($cryptSvc.Desc)) is DISABLED - RDP certificate operations will fail" "-FixRDPPermissions"
            }
        }

        # TLS 1.2 explicitly disabled in SCHANNEL
        foreach ($tlsRole in @('Client','Server')) {
            $tlsPath = "$ctrlRoot\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\$tlsRole"
            if (Test-Path $tlsPath) {
                $tlsProps = Get-ItemProperty $tlsPath -ErrorAction SilentlyContinue
                if ($tlsProps.Enabled -eq 0 -or $tlsProps.DisabledByDefault -eq 1) {
                    & $emit 'RDP' (& $toSev $sevRdpTlsDisabled) "TLS 1.2 $tlsRole is explicitly DISABLED in SCHANNEL - RDP SSL handshake may fail" "-FixRDP"
                }
            }
        }

        # NTLM restrictions
        $msv1 = Get-ItemProperty "$ctrlRoot\Lsa\MSV1_0" -ErrorAction SilentlyContinue
        if ($msv1) {
            if ($msv1.RestrictSendingNTLMTraffic -ge 2 -or $msv1.RestrictReceivingNTLMTraffic -ge 1) {
                & $emit 'RDP' (& $toSev $sevRdpNtlmRestrict) "NTLM restrictions: RestrictSending=$($msv1.RestrictSendingNTLMTraffic) RestrictReceiving=$($msv1.RestrictReceivingNTLMTraffic) - may block RDP auth" "-FixRDPAuth"
            }
        }

        # LmCompatibilityLevel (too restrictive blocks older clients)
        $lsaProps = Get-ItemProperty "$ctrlRoot\Lsa" -ErrorAction SilentlyContinue
        $lmCompat = $lsaProps.LmCompatibilityLevel
        if ($null -ne $lmCompat -and $lmCompat -gt 5) {
            & $emit 'RDP' (& $toSev $sevRdpLmCompat) "LmCompatibilityLevel=$lmCompat (>5) - may block NTLM-based RDP auth from some clients" "-FixRDPAuth"
        }

        # ── Credential Guard ────────────────────────────────────────────────
        $cgLsa = $lsaProps.LsaCfgFlags
        if ($cgLsa -and $cgLsa -ne 0) {
            $lockType = if ($cgLsa -eq 1) { 'UEFI lock - registry change alone is insufficient' } else { 'software lock' }
            & $emit 'Security' (& $toSev $sevCredentialGuard) "Credential Guard is enabled (LsaCfgFlags=$cgLsa, $lockType)" "-DisableCredentialGuard"
        }
        $runAsPPL = $lsaProps.RunAsPPL
        if ($runAsPPL -eq 1) {
            & $emit 'Security' (& $toSev $sevLsaPPL) 'LSA Protected Process (RunAsPPL=1) is active - may affect some security tools'
        }

        # ── Azure Guest Agent ────────────────────────────────────────────────
        foreach ($ag in @('WindowsAzureGuestAgent','RdAgent')) {
            $ap = "$svcRoot\$ag"
            if (Test-Path $ap) {
                $aStart = (Get-ItemProperty $ap -ErrorAction SilentlyContinue).Start
                if ($aStart -eq 4) {
                    & $emit 'AzureAgent' (& $toSev $sevAzureAgentDisabled) "$ag is DISABLED (Start=4) - VM will not respond to Azure platform operations" "-FixAzureGuestAgent"
                } elseif ($aStart -ne 2) {
                    & $emit 'AzureAgent' (& $toSev $sevAzureAgentWrongStart) "$ag Start=$aStart (expected 2/Auto)" "-FixAzureGuestAgent"
                } else {
                    & $emit 'AzureAgent' 'OK' "$ag Start=2 (Auto)"
                }
            } else {
                & $emit 'AzureAgent' (& $toSev $sevAzureAgentMissing) "$ag not found in registry - agent may not be installed" "-InstallAzureVMAgent"
            }
        }

        # ── Networking: BFE & TCP/IP ─────────────────────────────────────────
        $bfeStart = (Get-ItemProperty "$svcRoot\BFE" -ErrorAction SilentlyContinue).Start
        if ($bfeStart -eq 4) {
            & $emit 'Networking' (& $toSev $sevBfeDisabled) 'BFE (Base Filtering Engine) is DISABLED - Windows Firewall and IPSec/network policy will not function' "-EnableBFE"
        } elseif ($null -ne $bfeStart) {
            & $emit 'Networking' 'OK' "BFE Start=$bfeStart (enabled)"
        }
        $tcpStart = (Get-ItemProperty "$svcRoot\Tcpip" -ErrorAction SilentlyContinue).Start
        if ($tcpStart -eq 4) {
            & $emit 'Networking' (& $toSev $sevTcpipDisabled) 'Tcpip service is DISABLED - no network will be available' "-ResetNetworkStack"
        }

        # Additional networking services
        foreach ($netSvc in @(
            @{ N='Dnscache';         Desc='DNS Client - name resolution will fail' }
            @{ N='NlaSvc';           Desc='Network Location Awareness - network profile detection will fail' }
            @{ N='Dhcp';             Desc='DHCP Client - automatic IP configuration will not work' }
            @{ N='LanmanWorkstation'; Desc='Workstation service (SMB client) - file sharing access will fail' }
            @{ N='LanmanServer';      Desc='Server service (SMB server) - file sharing hosting will fail' }
        )) {
            $ns = (Get-ItemProperty "$svcRoot\$($netSvc.N)" -ErrorAction SilentlyContinue).Start
            if ($ns -eq 4) {
                & $emit 'Networking' (& $toSev $sevNetSvcDisabled) "$($netSvc.N) is DISABLED - $($netSvc.Desc)" "-ResetNetworkStack"
            }
        }

        # SAN policy (partmgr)
        $sanPolicy = (Get-ItemProperty "$svcRoot\partmgr\Parameters" -ErrorAction SilentlyContinue).SanPolicy
        if ($null -ne $sanPolicy -and $sanPolicy -ne 1 -and $sanPolicy -ne 0) {
            $sanDesc = switch ($sanPolicy) { 2 {'OfflineShared - shared disks stay offline'} 3 {'OfflineAll - all SAN disks stay offline'} 4 {'OfflineInternal - internal SAN disks offline'} default {"Value=$sanPolicy"} }
            & $emit 'Networking' (& $toSev $sevSanPolicy) "SAN policy is set to $sanDesc - data disks may not come online after migration; run -FixSanPolicy to set OnlineAll" "-FixSanPolicy"
        } elseif ($null -ne $sanPolicy) {
            & $emit 'Networking' 'OK' "SAN policy: OnlineAll ($sanPolicy)"
        }

        # ── Boot/System drivers with missing binaries ─────────────────────────
        # We only flag drivers that are registered to load at boot/system start
        # but whose binary is ABSENT on the offline disk - these WILL cause a
        # BSOD or hang on next boot. Inbox hardware drivers (LSI, Intel RST,
        # Broadcom HBA, etc.) ship with Windows but carry vendor company names;
        # they are NOT flagged here because they are safe and present on disk.
        # Use -DisableThirdPartyDrivers to intentionally suppress all non-MS
        # drivers when troubleshooting a clean-boot scenario.
        $missingDrivers = [System.Collections.Generic.List[string]]::new()
        Get-ChildItem $svcRoot -ErrorAction SilentlyContinue | ForEach-Object {
            $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
            # Only kernel/filesystem drivers (Type 1/2) at Boot(0) or System(1) start
            if ($p.Type -notin @(1,2) -or $p.Start -notin @(0,1) -or -not $p.ImagePath) { return }
            $imgR = $p.ImagePath `
                -replace '(?i)\\SystemRoot\\', "$script:WinDriveLetter\Windows\" `
                -replace '(?i)%SystemRoot%',   "$script:WinDriveLetter\Windows" `
                -replace '(?i)\\\?\?\\',      '' `
                -replace '(?i)^system32\\',   "$script:WinDriveLetter\Windows\System32\\"
            if ($imgR -match '^(.+?\.(?:sys|exe))') { $imgR = $Matches[1] }
            # Only flag if binary is genuinely missing
            if (-not (Test-Path $imgR)) { $missingDrivers.Add("$($_.PSChildName) ($($p.ImagePath))") }
        }
        if ($missingDrivers.Count -gt 0) {
            & $emit 'Drivers' (& $toSev $sevMissingDriverBinaries) "$($missingDrivers.Count) Boot/System driver(s) registered but binary MISSING - will BSOD on boot: $($missingDrivers -join ', ')" "-DisableThirdPartyDrivers"
        } else {
            & $emit 'Drivers' 'OK' 'All Boot/System drivers have binaries present on disk'
        }

        # ── Device class filters ─────────────────────────────────────────────
        $classRoot   = "HKLM:\BROKENSYSTEM\$csName\Control\Class"
        $safeFilters = @{
            '{4d36e967-e325-11ce-bfc1-08002be10318}' = [string[]]@('partmgr','fvevol','iorate','storqosflt','wcifs','ehstorclass')
            '{4d36e96a-e325-11ce-bfc1-08002be10318}' = [string[]]@('iasf','iastorf')
            '{4d36e97b-e325-11ce-bfc1-08002be10318}' = [string[]]@()
            '{71a27cdd-812a-11d0-bec7-08002be2092f}' = [string[]]@('volsnap','fvevol','rdyboost','spldr','volmgrx','iorate','storqosflt')
            '{4d36e972-e325-11ce-bfc1-08002be10318}' = [string[]]@('wfplwf','ndiscap','ndisimplatformmpfilter','vmsproxyhnicfilter','vms3cap','mslldp','psched','bridge')
        }
        $filterClasses = @(
            @{ GUID='{4d36e967-e325-11ce-bfc1-08002be10318}'; Name='DiskDrive';      Sev=(& $toSev $sevDeviceFiltersCrit) }
            @{ GUID='{4d36e96a-e325-11ce-bfc1-08002be10318}'; Name='SCSIAdapter';    Sev=(& $toSev $sevDeviceFiltersCrit) }
            @{ GUID='{4d36e97b-e325-11ce-bfc1-08002be10318}'; Name='SCSIController'; Sev=(& $toSev $sevDeviceFiltersCrit) }
            @{ GUID='{71a27cdd-812a-11d0-bec7-08002be2092f}'; Name='Volume';         Sev=(& $toSev $sevDeviceFiltersWarn) }
            @{ GUID='{4d36e972-e325-11ce-bfc1-08002be10318}'; Name='Net';            Sev=(& $toSev $sevDeviceFiltersWarn) }
        )
        foreach ($fc in $filterClasses) {
            $cp = "$classRoot\$($fc.GUID)"
            if (-not (Test-Path $cp)) { continue }
            $safe = $safeFilters[$fc.GUID]
            foreach ($ft in @('UpperFilters','LowerFilters')) {
                $raw     = (Get-ItemProperty $cp -ErrorAction SilentlyContinue).$ft
                $active  = @($raw | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() })
                $suspect = @($active | Where-Object { $safe -inotcontains $_ })
                if ($suspect.Count -gt 0) {
                    & $emit 'DeviceFilters' $fc.Sev "$($fc.Name) $ft contains non-standard entries: $($suspect -join ', ')" "-FixDeviceFilters"
                }
            }
        }

        # ── Orphaned NDIS bindings ────────────────────────────────────────────
        $orphanIds  = [System.Collections.Generic.List[string]]::new()
        $seenComp   = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($cg in @('{4D36E973-E325-11CE-BFC1-08002BE10318}','{4D36E974-E325-11CE-BFC1-08002BE10318}','{4D36E975-E325-11CE-BFC1-08002BE10318}')) {
            $ck = "HKLM:\BROKENSYSTEM\$csName\Control\Class\$cg"
            if (-not (Test-Path $ck)) { continue }
            Get-ChildItem $ck -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match '^\d{4}$' } | ForEach-Object {
                $pp  = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
                $cid = if ($pp.ComponentId) { $pp.ComponentId } else { $pp.ComponentID }
                if (-not $cid -or $cid -match '^ms_' -or -not $seenComp.Add($cid)) { return }
                $sn  = $cid -replace '^ms_',''
                $ipp = (Get-ItemProperty "$svcRoot\$sn" -ErrorAction SilentlyContinue).ImagePath
                if ($null -eq $ipp) { return }
                $ir  = $ipp -replace '(?i)\\SystemRoot\\',"$script:WinDriveLetter\Windows\" `
                           -replace '(?i)%SystemRoot%',  "$script:WinDriveLetter\Windows" `
                           -replace '(?i)\\\?\?\\',     '' `
                           -replace '(?i)^system32\\',  "$script:WinDriveLetter\Windows\System32\\"
                if ($ir -match '^(.+?\.(?:sys|exe|dll))') { $ir = $Matches[1] }
                if (-not (Test-Path $ir)) { $orphanIds.Add($cid) }
            }
        }
        if ($orphanIds.Count -gt 0) {
            & $emit 'Networking' (& $toSev $sevOrphanedNdis) "Orphaned NDIS binding(s) with missing binary: $($orphanIds -join ', ') - will prevent network initialisation at boot" "-FixNetBindings"
        } else {
            & $emit 'Networking' 'OK' 'No orphaned NDIS binding components'
        }

        # ── Windows Update services ──────────────────────────────────────────
        foreach ($wu in @('wuauserv','UsoSvc','WaaSMedicSvc')) {
            $wuS = (Get-ItemProperty "$svcRoot\$wu" -ErrorAction SilentlyContinue).Start
            if ($wuS -eq 4) {
                & $emit 'WindowsUpdate' (& $toSev $sevUpdateWuDisabled) "$wu is disabled (Start=4) - intentionally stopped via -DisableWindowsUpdate; re-enable on the live VM with: Set-Service -Name $wu -StartupType Automatic"
            }
        }

        # AppIDSvc (AppLocker enforcement service)
        $appIdSvc = (Get-ItemProperty "$svcRoot\AppIDSvc" -ErrorAction SilentlyContinue).Start
        if ($null -ne $appIdSvc -and $appIdSvc -ne 4) {
            & $emit 'Security' (& $toSev $sevAppIdSvc) "AppIDSvc (Application Identity) Start=$appIdSvc - this service is required for AppLocker to enforce rules"
        }

        # ── Hyper-V ACPI device IDs ─────────────────────────────────────────
        $enumAcpiRoot    = "HKLM:\BROKENSYSTEM\$csName\Enum\ACPI"
        $msft1000Present = Test-Path "$enumAcpiRoot\MSFT1000"
        $msft1002Present = Test-Path "$enumAcpiRoot\MSFT1002"
        if (-not $msft1000Present -or -not $msft1002Present) {
            $missing = @()
            if (-not $msft1000Present) { $missing += 'MSFT1000 (VMBus)' }
            if (-not $msft1002Present) { $missing += 'MSFT1002 (Hyper-V Gen Counter)' }
            & $emit 'ACPI' (& $toSev $sevACPISettings) "Hyper-V ACPI device $(if ($missing.Count -gt 1){'entries'}else{'entry'}) missing: $($missing -join ', ') - newer ACPI IDs needed for full Hyper-V synthetic device support" "-CopyACPISettings"
        } else {
            & $emit 'ACPI' 'OK' 'Hyper-V ACPI device entries present (MSFT1000, MSFT1002)'
        }
    }
    catch { Write-Warning "SYSTEM hive check error: $_" }
    finally { UnmountOffHive -Hive 'SYSTEM' }

    # ── 5. SOFTWARE hive ─────────────────────────────────────────────────────
    Write-Host "--- Registry (SOFTWARE hive)" -ForegroundColor DarkGray
    MountOffHive -WinPath $OfflineWindowsPath -Hive 'SOFTWARE'
    try {
        # OS edition & build
        $ntCv = Get-ItemProperty 'HKLM:\BROKENSOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
        if ($ntCv) {
            $build = if ($ntCv.CurrentBuildNumber) { "Build $($ntCv.CurrentBuildNumber).$($ntCv.UBR)" } else { '' }
            & $emit 'OS' 'INFO' "$($ntCv.ProductName)  |  Edition: $($ntCv.EditionID)  |  $build"
        }

        # CredSSP AllowEncryptionOracle
        $credSsp = (Get-ItemProperty 'HKLM:\BROKENSOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\CredSSP\Parameters' -ErrorAction SilentlyContinue).AllowEncryptionOracle
        if ($null -ne $credSsp -and $credSsp -ne 2) {
            $oDesc = switch ($credSsp) { 0 {'Force Updated Clients (most restrictive)'} 1 {'Mitigated'} default {"Value=$credSsp"} }
            & $emit 'RDP' (& $toSev $sevCredSspOracle) "CredSSP AllowEncryptionOracle=$credSsp ($oDesc) - may block RDP from clients without latest patches" "-FixRDPAuth"
        }

        # AppLocker enforced collections
        $srpBase = 'HKLM:\BROKENSOFTWARE\Policies\Microsoft\Windows\SrpV2'
        $enforced = @(foreach ($col in @('Exe','Dll','Script','Msi','Appx')) {
            $colPath = "$srpBase\$col"
            if ((Test-Path $colPath) -and (Get-ItemProperty $colPath -ErrorAction SilentlyContinue).EnforcementMode -ne 0) { $col }
        })
        if ($enforced.Count -gt 0) {
            & $emit 'Security' (& $toSev $sevAppLockerEnforcing) "AppLocker is enforcing rules for: $($enforced -join ', ') - may block processes from starting" "-DisableAppLocker"
        } else {
            & $emit 'Security' 'OK' 'AppLocker is not enforcing any rule collections'
        }

        # AppIDSvc - if enforcing above, check service state
        # (service state is in SYSTEM hive, checked there; this just notes correlation)

        # NLA policy via Group Policy / TS Policy path (SOFTWARE side)
        $tsPolicyBase = 'HKLM:\BROKENSOFTWARE\Policies\Microsoft\Windows NT\Terminal Services'
        if (Test-Path $tsPolicyBase) {
            $tsPol = Get-ItemProperty $tsPolicyBase -ErrorAction SilentlyContinue
            if ($null -ne $tsPol.fDenyTSConnections -and $tsPol.fDenyTSConnections -eq 1) {
                & $emit 'RDP' (& $toSev $sevGpRdpBlocked) 'Group Policy is BLOCKING RDP: Software\Policies\...\Terminal Services fDenyTSConnections=1 - -FixRDP clears this' "-FixRDP"
            }
            if ($null -ne $tsPol.UserAuthentication -and $tsPol.UserAuthentication -eq 0) {
                & $emit 'RDP' (& $toSev $sevGpNlaDisabled) 'Group Policy has disabled NLA (Software\Policies UserAuthentication=0) - this overrides the listener setting'
            }
            $sslFuncPath = 'HKLM:\BROKENSOFTWARE\Policies\Microsoft\Cryptography\Configuration\SSL\00010002'
            if (Test-Path $sslFuncPath) {
                $sslFunc = (Get-ItemProperty $sslFuncPath -ErrorAction SilentlyContinue).Functions
                if ($null -ne $sslFunc) {
                    & $emit 'RDP' (& $toSev $sevSslCipherPolicy) "SSL cipher suite policy (SSL\00010002 Functions) is configured - may restrict TLS cipher suites available to RDP; run -FixRDP to clear" "-FixRDP"
                }
            }
        }

        # CBS / Component Based Servicing pending state
        $cbsBase = 'HKLM:\BROKENSOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing'
        $cbsIssues = [System.Collections.Generic.List[string]]::new()
        foreach ($cbsKey in @('RebootPending','PackagesPending','SessionsPending')) {
            $cbsKeyPath = "$cbsBase\$cbsKey"
            if (-not (Test-Path $cbsKeyPath)) { continue }

            # Collect detail: subkey names + any values on the key itself
            $subkeys = @(Get-ChildItem $cbsKeyPath -ErrorAction SilentlyContinue | Select-Object -ExpandProperty PSChildName)
            $vals    = Get-ItemProperty $cbsKeyPath -ErrorAction SilentlyContinue
            $valNames = @($vals.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | Select-Object -ExpandProperty Name)

            $detail = switch ($cbsKey) {
                'RebootPending' {
                    # Values here are package names pending reboot
                    if ($valNames.Count -gt 0) { "pending package(s): $($valNames[0..([Math]::Min(2,$valNames.Count-1))] -join ', ')$(if ($valNames.Count -gt 3){ " (+$($valNames.Count-3) more)" })" }
                    else { 'key exists (no values)' }
                }
                'PackagesPending' {
                    # Subkeys are package names
                    if ($subkeys.Count -gt 0) { "$($subkeys.Count) package(s) pending: $($subkeys[0..([Math]::Min(1,$subkeys.Count-1))] -join ', ')$(if ($subkeys.Count -gt 2){ " (+$($subkeys.Count-2) more)" })" }
                    else { 'key exists (no subkeys)' }
                }
                'SessionsPending' {
                    # Each subkey is a CBS session; the Exclusive value indicates a locked session
                    $exclusive = @($subkeys | Where-Object {
                        $sv = (Get-ItemProperty "$cbsKeyPath\$_" -ErrorAction SilentlyContinue).Exclusive
                        $null -ne $sv -and $sv -ne 0
                    })
                    if ($exclusive.Count -gt 0) {
                        "EXCLUSIVE session lock(s) present - a CBS operation was interrupted mid-flight; session ID(s): $($exclusive -join ', ')"
                    } elseif ($subkeys.Count -gt 0) {
                        "$($subkeys.Count) session entry(s) - likely a previous CBS transaction that did not complete cleanly; session ID(s): $($subkeys -join ', ')"
                    } else {
                        'key exists (no subkeys) - stale CBS session marker'
                    }
                }
            }
            $cbsIssues.Add("$cbsKey ($detail)")
        }
        if ($cbsIssues.Count -gt 0) {
            foreach ($issue in $cbsIssues) {
                # SessionsPending without Exclusive locks is normal on healthy systems — emit INFO only
                $isSessionInfo = ($issue -match '^SessionsPending' -and $issue -notmatch 'EXCLUSIVE')
                $level = if ($isSessionInfo) { (& $toSev $sevCbsPendingInfo) } else { (& $toSev $sevCbsPendingWarn) }
                $suffix = if ($isSessionInfo) { ' (stale entries, no exclusive lock - normal on healthy systems)' } else { " - may cause 'Configuring Windows Updates' boot loop" }
                $fix   = if ($isSessionInfo) { $null } else { '-FixPendingUpdates' }
                & $emit 'WindowsUpdate' $level "CBS pending state: $issue$suffix" $fix
            }
        } else {
            & $emit 'WindowsUpdate' 'OK' 'No CBS pending state keys detected'
        }
    }
    catch { Write-Warning "SOFTWARE hive check error: $_" }
    finally { UnmountOffHive -Hive 'SOFTWARE' }

    # ── 6. Azure Agent binaries on disk ──────────────────────────────────────
    Write-Host "--- Azure VM Agent" -ForegroundColor DarkGray
    $azureDir = Join-Path $script:WinDriveLetter 'WindowsAzure'
    if (Test-Path $azureDir) {
        $gaDirs = @(Get-ChildItem $azureDir -Filter 'GuestAgent_*' -Directory -ErrorAction SilentlyContinue)
        if ($gaDirs.Count -gt 0) {
            & $emit 'AzureAgent' 'OK' "Agent binaries present: $($gaDirs[-1].Name)"
        } else {
            & $emit 'AzureAgent' (& $toSev $sevAzureAgentMissing) 'WindowsAzure folder exists but no GuestAgent_* subfolder found' "-InstallAzureVMAgent"
        }
    } else {
        & $emit 'AzureAgent' (& $toSev $sevAzureAgentMissing) "WindowsAzure folder not found on $script:WinDriveLetter - VM Agent may not be installed" "-InstallAzureVMAgent"
    }

    # ── 7. RDP private key / MachineKeys ─────────────────────────────────────
    Write-Host "--- RDP Certificate & MachineKeys" -ForegroundColor DarkGray
    $machineKeysPath = Join-Path $script:WinDriveLetter 'ProgramData\Microsoft\Crypto\RSA\MachineKeys'
    if (Test-Path $machineKeysPath) {
        # The RDP private key file has a well-known name prefix f686aace6942fb7f7ceb231212eef4a4
        $rdpKeyFiles = @(Get-ChildItem $machineKeysPath -Filter 'f686aace6942fb7f7ceb231212eef4a4*' -ErrorAction SilentlyContinue)
        if ($rdpKeyFiles.Count -gt 0) {
            & $emit 'RDP' 'OK' "RDP private key file present ($($rdpKeyFiles[0].Name))"

            # Check ACLs on the key file. Required principals (matched by well-known SID):
            #   S-1-5-18  = NT AUTHORITY\SYSTEM        -> needs FullControl
            #   S-1-5-20  = NT AUTHORITY\NETWORK SERVICE -> needs at least Read (TermService runs as this)
            #   S-1-5-80-* matching SessionEnv service  -> needs FullControl (checked by display name)
            try {
                $keyAcl = Get-Acl -Path $rdpKeyFiles[0].FullName -ErrorAction Stop
                $rules  = $keyAcl.Access

                $sidSystem  = [System.Security.Principal.SecurityIdentifier]'S-1-5-18'
                $sidNetSvc  = [System.Security.Principal.SecurityIdentifier]'S-1-5-20'

                $systemOk  = $rules | Where-Object {
                    try { $_.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]) -eq $sidSystem } catch { $false }
                } | Where-Object { $_.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::FullControl }

                $netSvcOk  = $rules | Where-Object {
                    try { $_.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]) -eq $sidNetSvc } catch { $false }
                } | Where-Object { $_.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::Read }

                # SessionEnv runs as NT Service\SessionEnv - no fixed SID, match by identity string
                $sessionEnvOk = $rules | Where-Object {
                    $_.IdentityReference.Value -match 'SessionEnv'
                } | Where-Object { $_.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::FullControl }

                if (-not $systemOk) {
                    & $emit 'RDP' (& $toSev $sevRdpKeySystemAcl) "RDP private key: NT AUTHORITY\SYSTEM does not have FullControl - RDP service may fail to use the certificate" "-FixRDPPermissions"
                }
                if (-not $netSvcOk) {
                    & $emit 'RDP' (& $toSev $sevRdpKeyNetSvcAcl) "RDP private key: NT AUTHORITY\NETWORK SERVICE does not have Read access - TermService will fail to load the RDP certificate" "-FixRDPPermissions"
                } else {
                    & $emit 'RDP' 'OK' 'RDP private key ACLs: SYSTEM and NETWORK SERVICE have required permissions'
                }
                if (-not $sessionEnvOk) {
                    & $emit 'RDP' (& $toSev $sevRdpKeySessionEnvAcl) 'RDP private key: NT Service\SessionEnv FullControl not found - may cause RDP session issues on some Windows versions' "-FixRDPPermissions"
                }
            } catch {
                & $emit 'RDP' 'INFO' "Could not read ACLs on RDP private key file: $_"
            }
        } else {
            & $emit 'RDP' (& $toSev $sevRdpKeyFileMissing) 'RDP private key file (f686aace...) not found in MachineKeys - RDP certificate will need to be regenerated' "-FixRDPCert"
        }
        # Check for zero-length or suspiciously small key files (corrupted)
        $emptyKeys = @(Get-ChildItem $machineKeysPath -ErrorAction SilentlyContinue | Where-Object { $_.Length -eq 0 })
        if ($emptyKeys.Count -gt 0) {
            & $emit 'RDP' (& $toSev $sevRdpKeyZeroLength) "$($emptyKeys.Count) zero-length file(s) found in MachineKeys - may indicate corrupted key store; run -FixRDPPermissions" "-FixRDPPermissions"
        }
    } else {
        & $emit 'RDP' (& $toSev $sevMachineKeysMissing) 'MachineKeys folder missing - RDP certificate operations will fail on boot' "-FixRDPPermissions"
    }

    # ── Summary ──────────────────────────────────────────────────────────────
    $crits = @($findings | Where-Object Severity -eq 'CRIT')
    $warns = @($findings | Where-Object Severity -eq 'WARN')
    $infos = @($findings | Where-Object Severity -eq 'INFO')
    $oks   = @($findings | Where-Object Severity -eq 'OK')

    Write-Host "`n===================================================================" -ForegroundColor Cyan
    Write-Host "  System Check Summary  -  $($crits.Count) critical  /  $($warns.Count) warnings  /  $($infos.Count) info  /  $($oks.Count) ok" -ForegroundColor Cyan
    Write-Host "===================================================================" -ForegroundColor Cyan

    if ($crits.Count -gt 0) {
        Write-Host "`nCRITICAL - likely preventing boot or connectivity:" -ForegroundColor Red
        foreach ($f in $crits) {
            Write-Host "  [$($f.Category)] $($f.Description)" -ForegroundColor Red
            if ($f.Fix) { Write-Host "    > .\$scriptFile -DiskNumber $script:DiskNumber $($f.Fix)" -ForegroundColor DarkCyan }
        }
    }
    if ($warns.Count -gt 0) {
        Write-Host "`nWARNINGS - should be investigated:" -ForegroundColor Yellow
        foreach ($f in $warns) {
            Write-Host "  [$($f.Category)] $($f.Description)" -ForegroundColor Yellow
            if ($f.Fix) { Write-Host "    > .\$scriptFile -DiskNumber $script:DiskNumber $($f.Fix)" -ForegroundColor DarkCyan }
        }
    }
    if ($crits.Count -eq 0 -and $warns.Count -eq 0) {
        Write-Host "`nNo critical issues or warnings detected." -ForegroundColor Green
    }

    Write-Host ""
    Write-ActionLog -Event 'RunSystemCheck' -Details @{
        Critical = $crits.Count
        Warnings = $warns.Count
        Findings = ($findings | ForEach-Object { "$($_.Severity)[$($_.Category)] $($_.Description)" }) -join ' | '
    }
}

function ResetNetworkStack {
    Write-Host "Placing network stack reset script to run at next boot..." -ForegroundColor Yellow
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        Set-ItemProperty-Logged -Path "HKLM:\BROKENSYSTEM\Setup" -Name SetupType -Value 2 -Type DWord -Force
        Set-ItemProperty-Logged -Path "HKLM:\BROKENSYSTEM\Setup" -Name CmdLine -Value "cmd.exe /c c:\temp\resetnet.cmd" -Type String -Force

        $netScript = @"
@echo off
netsh int ip reset C:\temp\netsh_ip_reset.log
netsh winsock reset
netsh advfirewall reset
rem Re-enable critical inbound rules that advfirewall reset disables by default
netsh advfirewall firewall set rule group="remote desktop" new enable=yes
netsh advfirewall firewall set rule group="file and printer sharing" new enable=yes
netsh advfirewall firewall set rule group="windows remote management" new enable=yes
ipconfig /flushdns
reg add "HKLM\SYSTEM\Setup" /v SetupType /t REG_DWORD /d 0 /f
reg add "HKLM\SYSTEM\Setup" /v CmdLine /t REG_SZ /d "" /f
del /F C:\Temp\resetnet.cmd > NUL
shutdown /r /t 10 /c "Network stack reset complete - rebooting"
"@
        Ensure-GuestTempDir
        New-Item-Logged -Path "$script:WinDriveLetter\Temp" -Name "resetnet.cmd" -ItemType File -Value $netScript.Trim() -Force

        Write-Host "Network stack reset script placed. On next boot: IP stack, Winsock and firewall will be reset, then VM will reboot automatically." -ForegroundColor Green
    }
    catch {
        Write-Error "ResetNetworkStack failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

function DisableWindowsUpdate {
    Write-Host "Disabling Windows Update services to prevent boot loops from update processing..." -ForegroundColor Yellow
    $OfflineWindowsPath = Join-Path $script:WinDriveLetter "Windows"
    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        $SystemRoot = Get-SystemRootPath

        $services = @('wuauserv', 'UsoSvc', 'WaaSMedicSvc', 'UpdateOrchestrator')
        foreach ($svc in $services) {
            $svcPath = "$SystemRoot\Services\$svc"
            if (Test-Path $svcPath) {
                $before = (Get-ItemProperty $svcPath -ErrorAction SilentlyContinue).Start
                Set-ItemProperty-Logged -Path $svcPath -Name Start -Value 4 -Type DWord -Force
                Write-Host "  $svc disabled (was $before)" -ForegroundColor Yellow
            }
        }

        Write-Host "Windows Update services disabled. Re-enable them after recovery." -ForegroundColor Green
        Write-Host "`n--- REVERT COMMANDS (run on the VM after recovery) ---" -ForegroundColor Cyan
        Write-Host "Set-Service wuauserv -StartupType Automatic" -ForegroundColor White
        Write-Host "Set-Service UsoSvc   -StartupType Automatic" -ForegroundColor White
        Write-Host "# Or use: sc.exe config wuauserv start= auto" -ForegroundColor DarkGray
        Write-Host "------------------------------------------------------`n" -ForegroundColor Cyan
    }
    catch {
        Write-Error "DisableWindowsUpdate failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

function RestoreRegistryFromRegBack {
    if (-not (Confirm-CriticalOperation -Operation 'Restore Registry from RegBack (-RestoreRegistryFromRegBack)' -Details @"
Replaces the live SYSTEM and SOFTWARE hives with the RegBack copies.
The current hives are renamed to *.bak before overwriting.
"@)) { return }

    Write-Host "Attempting to restore SYSTEM and SOFTWARE hives from RegBack..." -ForegroundColor Yellow
    $regBackPath = Join-Path $script:WinDriveLetter "Windows\System32\config\RegBack"
    $configPath  = Join-Path $script:WinDriveLetter "Windows\System32\config"

    if (-not (Test-Path $regBackPath)) {
        Write-Warning "RegBack folder not found at $regBackPath. Cannot restore."
        return
    }

    $hives = @('SYSTEM', 'SOFTWARE')
    foreach ($hive in $hives) {
        $backupFile = Join-Path $regBackPath $hive
        $liveFile   = Join-Path $configPath  $hive

        if (-not (Test-Path $backupFile)) {
            Write-Warning "  RegBack copy of $hive not found - skipping."
            continue
        }

        $backupSize = (Get-Item $backupFile).Length
        if ($backupSize -lt 1MB) {
            Write-Warning "  RegBack $hive is suspiciously small ($backupSize bytes) - skipping to avoid overwriting with empty hive."
            continue
        }

        Write-Host "  Restoring $hive ($([math]::Round($backupSize/1MB,1)) MB)..." -ForegroundColor Yellow
        $bakPath = New-UniqueBackupPath -BasePath $liveFile -BakSuffix ".bak"
        Move-Item-Logged -LiteralPath $liveFile -Destination $bakPath -Force
        Copy-Item-Logged -Path $backupFile -Destination $liveFile -Force
        Write-Host "  $hive restored. Previous live hive backed up to $bakPath" -ForegroundColor Green
    }
    Write-Host "RegBack restore complete. Boot the VM to verify." -ForegroundColor Green
}

function SetTestSigning {
    param([bool]$Enable)
    $action = if ($Enable) { "Enabling" } else { "Disabling" }
    Write-Host "$action test signing mode..." -ForegroundColor Yellow
    try {
        $storePath = Get-BcdStorePath -Generation $script:VMGen -BootDrive $script:BootDriveLetter
        $identifier = Get-BcdBootLoaderId -StorePath $storePath
        if (-not $identifier) { return }

        $val = if ($Enable) { "on" } else { "off" }
        Invoke-BcdEdit -StorePath $storePath -Command "/set $identifier testsigning $val"

        if ($Enable) {
            Write-Warning "Test signing is now ENABLED. This reduces security by allowing unsigned drivers. Disable after recovery with -DisableTestSigning."
        }
        else {
            Write-Host "Test signing disabled." -ForegroundColor Green
        }
    }
    catch {
        Write-Error "SetTestSigning failed: $_"
        throw
    }
}

function CheckDiskHealth {
    Write-Host "`nDisk Health Report" -ForegroundColor Cyan
    Write-Host "==================" -ForegroundColor Cyan

    $targetDisk = Get-Disk -Number $script:DiskNumber -ErrorAction SilentlyContinue
    if (-not $targetDisk) { Write-Warning "Could not retrieve disk info."; return }

    Write-Host "Disk $script:DiskNumber: $($targetDisk.FriendlyName)"
    Write-Host "  Style      : $($targetDisk.PartitionStyle)"
    Write-Host "  Size       : $([math]::Round($targetDisk.Size/1GB,2)) GB"
    Write-Host "  Health     : $($targetDisk.HealthStatus)"
    Write-Host "  Operational: $($targetDisk.OperationalStatus)"

    $partitions = Get-Partition -DiskNumber $script:DiskNumber -ErrorAction SilentlyContinue
    Write-Host "`nPartitions:" -ForegroundColor Cyan
    foreach ($p in $partitions) {
        $sizeMB = [math]::Round($p.Size / 1MB)
        $letter = if ($p.DriveLetter) { "$($p.DriveLetter):" } else { "(no letter)" }
        $vol    = Get-Volume -Partition $p -ErrorAction SilentlyContinue
        $fs     = if ($vol) { $vol.FileSystemType } else { "N/A" }
        $health = if ($vol) { $vol.HealthStatus } else { "N/A" }
        $raw    = if ($fs -eq 'Unknown' -or $fs -eq '') { " !!! RAW FILESYSTEM - data may be inaccessible" } else { "" }

        $color  = if ($raw) { 'Red' } elseif ($health -ne 'Healthy' -and $health -ne 'N/A') { 'Yellow' } else { 'White' }
        Write-Host ("  Partition {0,2}: {1,-10} | {2,-12} | {3} MB | FS: {4,-8} | Health: {5}{6}" -f `
            $p.PartitionNumber, $letter, $p.Type, $sizeMB, $fs, $health, $raw) -ForegroundColor $color
    }
    Write-Host ""
    Write-Host "Windows drive : $script:WinDriveLetter"  -ForegroundColor Green
    Write-Host "Boot drive    : $script:BootDriveLetter" -ForegroundColor Green
    Write-Host "VM Generation : Gen$script:VMGen"        -ForegroundColor Green
    Write-Host ""

    Invoke-Logged -Description "CheckDiskHealth" -Details @{ Disk = $script:DiskNumber; PartitionCount = $partitions.Count } -ScriptBlock { "OK" } | Out-Null
}

function CollectEventLogs {
    $destBase = if (Test-Path "C:\temp") { "C:\temp" } else { New-Item "C:\temp" -ItemType Directory -Force | Select-Object -ExpandProperty FullName }
    $destFolder = Join-Path $destBase "GuestEventLogs_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    New-Item -Path $destFolder -ItemType Directory -Force | Out-Null

    Write-Host "Collecting guest event logs to $destFolder ..." -ForegroundColor Yellow

    $logSources = @(
        @{ Src = "$script:WinDriveLetter\Windows\System32\winevt\Logs\System.evtx";      Dst = "System.evtx" },
        @{ Src = "$script:WinDriveLetter\Windows\System32\winevt\Logs\Application.evtx"; Dst = "Application.evtx" },
        @{ Src = "$script:WinDriveLetter\Windows\System32\winevt\Logs\Security.evtx";    Dst = "Security.evtx" },
        @{ Src = "$script:WinDriveLetter\Windows\inf\setupapi.dev.log";                  Dst = "setupapi.dev.log" },
        @{ Src = "$script:WinDriveLetter\Windows\inf\setupapi.setup.log";                Dst = "setupapi.setup.log" },
        @{ Src = "$script:WinDriveLetter\Windows\ntbtlog.txt";                           Dst = "ntbtlog.txt" },
        @{ Src = "$script:WinDriveLetter\Windows\Minidump";                              Dst = "Minidump" }
    )

    foreach ($log in $logSources) {
        if (Test-Path $log.Src) {
            $dstPath = Join-Path $destFolder $log.Dst
            Copy-Item-Logged -Path $log.Src -Destination $dstPath -Recurse -Force
            Write-Host "  Copied: $($log.Dst)" -ForegroundColor Green
        }
        else {
            Write-Host "  Not found (skipped): $($log.Dst)" -ForegroundColor DarkGray
        }
    }

    Write-Host "Event logs collected to: $destFolder" -ForegroundColor Green
    Write-Host "Tip: Use 'Get-WinEvent -Path <.evtx>' or Event Viewer to open .evtx files from the host." -ForegroundColor DarkCyan
}

function ResetUserRights {
    $OfflineWindowsPath = Join-Path $WinDriveLetter "Windows"

    MountOffHive -WinPath $OfflineWindowsPath -Hive "SYSTEM"
    try {
        Write-Host "Configuring setup and adding script..." -ForegroundColor Green
        Set-ItemProperty-Logged -Path "HKLM:\BROKENSYSTEM\Setup" -Name SetupType -Value 2 -Type DWord -Force
        Set-ItemProperty-Logged -Path "HKLM:\BROKENSYSTEM\Setup" -Name CmdLine -Value "cmd.exe /c c:\temp\resetuserrights.cmd" -Type String -Force

        $resetuserrights = @"
@echo off
secedit /export /areas USER_RIGHTS /cfg C:\temp\UserRightsBefore.txt
secedit /configure /cfg C:\Windows\inf\defltbase.inf /db defltbase.sdb /verbose
secedit /export /areas USER_RIGHTS /cfg C:\temp\UserRightsAfter.txt
reg add "HKLM\SYSTEM\Setup" /v SetupType /t REG_DWORD /d 0 /f
reg add "HKLM\SYSTEM\Setup" /v CmdLine /t REG_SZ /d "" /f
del /F C:\temp\resetuserrights.cmd > NUL
"@

        Ensure-GuestTempDir
        New-Item-Logged -Path "$WinDriveLetter\Temp" -Name "resetuserrights.cmd" -ItemType File -Value $resetuserrights.ToString().Trim() -Force
    }
    catch {
        Write-Error "ResetUserRights failed: $_"
        throw
    }
    finally {
        UnmountOffHive -Hive "SYSTEM"
    }
}

################################################################################
# Initialize-TargetDisk
#
# Resolves a VM name (Hyper-V) or a raw disk number to a physical disk,
# brings it online, assigns temporary drive letters to unmounted partitions,
# detects the Windows and Boot partitions, and populates the script-scoped
# variables consumed by all repair functions:
#   $script:WinDriveLetter   – e.g. "E:\"
#   $script:BootDriveLetter  – e.g. "F:\" (may equal WinDriveLetter on Gen1)
#   $script:VMGen            – 1 (MBR/BIOS) or 2 (GPT/UEFI)
#
# Returns $true on success, $false on failure.
function Initialize-TargetDisk {
    param(
        [string]$VMName = "",
        [int]   $DiskNumber = -1
    )

    # ------------------------------------------------------------------
    # 1. Resolve VMName -> DiskNumber (Hyper-V path)
    # ------------------------------------------------------------------
    if (-not [string]::IsNullOrWhiteSpace($VMName)) {
        Write-Host "Resolving Hyper-V VM '$VMName'..." -ForegroundColor Cyan

        $vm = Get-VM -Name $VMName -ErrorAction SilentlyContinue
        if (-not $vm) {
            Write-Error "VM '$VMName' not found."
            return $false
        }

        # Stop VM before accessing its disk
        if ($vm.State -ne 'Off') {
            Write-Host "Stopping VM '$VMName'..."
            $script:VMIPAddress = (Get-VM -Name $VMName | Get-VMNetworkAdapter).IPAddresses |
            Where-Object { $_ -match "\." } | Select-Object -First 1
            $vm | Stop-VM -TurnOff -Force
            Start-Sleep -Seconds 2
        }

        # Primary: use controller 0:0 DiskNumber (Hyper-V populates this when VHD is attached)
        $osDrive = $vm.HardDrives | Where-Object { $_.ControllerNumber -eq 0 -and $_.ControllerLocation -eq 0 }
        $DiskNumber = if ($osDrive) { $osDrive.DiskNumber } else { -1 }

        # Fallback: Get-VHD exposes DiskNumber when the VHD is mounted as a loopback disk
        if ($null -eq $DiskNumber -or $DiskNumber -lt 0) {
            foreach ($vhdDrive in ($vm | Get-VMHardDiskDrive)) {
                $vhd = Get-VHD -Path $vhdDrive.Path -ErrorAction SilentlyContinue
                if ($vhd -and $null -ne $vhd.DiskNumber -and $vhd.DiskNumber -ge 0) {
                    $DiskNumber = $vhd.DiskNumber
                    Write-Host "Matched disk $DiskNumber via Get-VHD: $($vhdDrive.Path)" -ForegroundColor Green
                    break
                }
            }
        }

        if ($null -eq $DiskNumber -or $DiskNumber -lt 0) {
            Write-Error "Could not determine disk number for VM '$VMName'. Attach the disk manually and use -DiskNumber instead."
            return $false
        }

        Write-Host "[OK] VM disk resolved to Disk $DiskNumber`n" -ForegroundColor Green
    }

    if ($DiskNumber -lt 0) {
        Write-Error "No disk specified. Provide -VMName (Hyper-V) or -DiskNumber."
        return $false
    }

    # Guard: never operate on the local OS disk - applies to both -DiskNumber and -VMName paths
    if ($null -ne $script:LocalOsDiskNumber -and $DiskNumber -eq $script:LocalOsDiskNumber) {
        Write-Error "Disk $DiskNumber hosts the local C: drive. Cannot use it as a repair target - this would risk breaking the Hyper-V host."
        return $false
    }

    # ------------------------------------------------------------------
    # 2. Get disk and validate
    # ------------------------------------------------------------------
    $targetDisk = Get-Disk -Number $DiskNumber -ErrorAction SilentlyContinue
    if (-not $targetDisk) {
        Write-Error "Disk $DiskNumber not found. Use the script without parameters to list available disks."
        return $false
    }

    Write-Host "Target Disk: Disk $DiskNumber | $($targetDisk.PartitionStyle) | $([math]::Round($targetDisk.Size / 1GB)) GB"

    # ------------------------------------------------------------------
    # 2.5 Check if disk is currently owned by a running VM
    # ------------------------------------------------------------------
    $owningVM = if (Get-Command Get-VM -ErrorAction SilentlyContinue) {
        Get-VM -ErrorAction SilentlyContinue | Where-Object { $_.State -ne 'Off' } | ForEach-Object {
            $checkVM = $_
            foreach ($hd in $checkVM.HardDrives) {
                $resolvedNum = -1
                if ($null -ne $hd.DiskNumber -and [int]$hd.DiskNumber -ge 0) {
                    $resolvedNum = [int]$hd.DiskNumber
                }
                elseif ($hd.Path -match '\.vhdx?$') {
                    $vhd = Get-VHD -Path $hd.Path -ErrorAction SilentlyContinue
                    if ($vhd -and $null -ne $vhd.DiskNumber -and [int]$vhd.DiskNumber -ge 0) {
                        $resolvedNum = [int]$vhd.DiskNumber
                    }
                }
                if ($resolvedNum -eq $DiskNumber) {
                    $checkVM.Name
                    return
                }
            }
        } | Select-Object -First 1
    }
    if ($owningVM) {
        Write-Error "Disk $DiskNumber is currently attached to running VM '$owningVM'. Stop the VM first, or re-run with -VMName '$owningVM' to have the script stop it automatically."
        return $false
    }

    # ------------------------------------------------------------------
    # 3. Bring disk online and writable
    # ------------------------------------------------------------------
    $usedLetters = (Get-Volume).DriveLetter | Where-Object { $_ } | ForEach-Object { ([string]$_).ToUpper() }

    if ($targetDisk.OperationalStatus -eq 'Offline' -or $targetDisk.IsOffline) {
        Write-Host "Bringing disk $DiskNumber online..." -ForegroundColor Cyan
        try {
            $targetDisk | Set-Disk -IsOffline $false
            $targetDisk | Set-Disk -IsReadOnly $false
        }
        catch {
            if ($_ -match 'Access Denied|40001') {
                Write-Error "Access denied bringing disk $DiskNumber online. The disk is likely in use by a running VM. Stop the VM first, or re-run with -VMName <name> to have the script stop it automatically."
            }
            else {
                Write-Error "Failed to bring disk $DiskNumber online: $($_.Exception.Message)"
            }
            return $false
        }
        Start-Sleep -Seconds 2
        $targetDisk = Get-Disk -Number $DiskNumber
    }
    elseif ($targetDisk.IsReadOnly) {
        Write-Host "Clearing read-only flag on disk $DiskNumber..." -ForegroundColor Cyan
        try {
            $targetDisk | Set-Disk -IsReadOnly $false
        }
        catch {
            if ($_ -match 'Access Denied|40001') {
                Write-Error "Access denied clearing read-only flag on disk $DiskNumber. The disk is likely in use by a running VM."
            }
            else {
                Write-Error "Failed to clear read-only flag on disk ${DiskNumber}: $($_.Exception.Message)"
            }
            return $false
        }
        $targetDisk = Get-Disk -Number $DiskNumber
    }

    # ------------------------------------------------------------------
    # 4. Detect generation (MBR = Gen1, GPT = Gen2)
    # ------------------------------------------------------------------
    $script:VMGen = Get-DiskGeneration -Disk $targetDisk
    Write-Host "Detected Partition Style: $(if ($script:VMGen -eq 1) { 'MBR (BIOS) - Gen1' } else { 'GPT (UEFI) - Gen2' })"

    # ------------------------------------------------------------------
    # 5. Assign drive letters to unmounted useful partitions
    # ------------------------------------------------------------------
    $partitions = $targetDisk | Get-Partition
    foreach ($part in $partitions) {
        if (-not $part.DriveLetter -and ($part.Type -in @('System', 'Basic', 'IFS'))) {
            Write-Host "Assigning drive letter to Partition $($part.PartitionNumber) ($($part.Type))..."
            $part | Add-PartitionAccessPath -AssignDriveLetter -ErrorAction SilentlyContinue
        }
    }

    # ------------------------------------------------------------------
    # 6. Detect Windows and Boot partitions via AccessPaths
    # ------------------------------------------------------------------
    $winDrive = $null
    $bootDrive = $null

    $partitions = $targetDisk | Get-Partition
    Write-Host "`nPartition Analysis:" -ForegroundColor Cyan
    foreach ($part in $partitions) {
        $accessPaths = $part.AccessPaths | Where-Object { $_ -match ":" }
        foreach ($accessPath in $accessPaths) {
            $role = Get-PartitionRole -AccessPath $accessPath -PartitionInfo $part
            Write-Host "  Partition $($part.PartitionNumber) ($accessPath): $role"

            if ((Test-Path (Join-Path $accessPath "Windows\System32\ntdll.dll")) -and -not $winDrive) {
                $winDrive = $accessPath
            }

            if ($script:VMGen -eq 1) {
                # Active MBR partition with Boot\BCD present
                if ($part.IsActive -and (Test-Path (Join-Path $accessPath "Boot\BCD")) -and -not $bootDrive) {
                    $bootDrive = $accessPath
                }
                # Fallback: active MBR partition even without BCD (BCD missing/corrupted)
                elseif ($part.IsActive -and -not $bootDrive) {
                    $bootDrive = $accessPath
                    Write-Host "  ** Boot partition identified by MBR Active flag (BCD file missing - run -FixBoot to rebuild)" -ForegroundColor Yellow
                }
            }
            elseif ($script:VMGen -eq 2) {
                # GPT System partition with EFI\Microsoft\Boot\BCD present
                if ($part.Type -eq 'System' -and (Test-Path (Join-Path $accessPath "EFI\Microsoft\Boot\BCD")) -and -not $bootDrive) {
                    $bootDrive = $accessPath
                }
                # Fallback 1: GPT System partition with EFI\Microsoft\Boot folder (BCD deleted)
                elseif ($part.Type -eq 'System' -and (Test-Path (Join-Path $accessPath "EFI\Microsoft\Boot")) -and -not $bootDrive) {
                    $bootDrive = $accessPath
                    Write-Host "  ** Boot partition identified by GPT System type + EFI folder (BCD missing - run -FixBoot to rebuild)" -ForegroundColor Yellow
                }
                # Fallback 2: GPT System partition exists but EFI folder is also gone
                elseif ($part.Type -eq 'System' -and -not $bootDrive) {
                    $bootDrive = $accessPath
                    Write-Host "  ** Boot partition identified by GPT System partition type only (EFI/BCD missing - run -FixBoot to rebuild)" -ForegroundColor Yellow
                }
            }
        }
    }

    $newUsedLetters = (Get-Volume).DriveLetter | Where-Object { $_ } | ForEach-Object { ([string]$_).ToUpper() }
    $addedLetters = $newUsedLetters | Where-Object { $usedLetters -notcontains $_ }
    if ($addedLetters) { Write-Host "Added letter(s): $($addedLetters -join ', ')" }

    # ------------------------------------------------------------------
    # 7. Validate and publish script-scoped variables
    # ------------------------------------------------------------------
    if (-not $winDrive) {
        Write-Error "Cannot proceed: Windows partition (containing \Windows\System32\ntdll.dll) not detected on disk $DiskNumber."
        return $false
    }

    if (-not $bootDrive) {
        Write-Host "No separate boot partition found; using Windows drive as boot drive." -ForegroundColor Yellow
        $bootDrive = $winDrive
    }

    # Normalize to trailing backslash
    if ($winDrive -notmatch '\\$') { $winDrive = "$winDrive\" }
    if ($bootDrive -notmatch '\\$') { $bootDrive = "$bootDrive\" }

    $script:WinDriveLetter = $winDrive
    $script:BootDriveLetter = $bootDrive
    $script:DiskNumber = $DiskNumber

    Write-Host ""
    Write-Host "[OK] Windows drive : $script:WinDriveLetter"  -ForegroundColor Green
    Write-Host "[OK] Boot drive    : $script:BootDriveLetter" -ForegroundColor Green
    Write-Host "[OK] VM Generation : Gen$script:VMGen"        -ForegroundColor Green
    Write-Host ""

    return $true
}

################################################################################
# Consolidated helper: Repair-OfflineDisk
#
# Parameters: same as the script-level parameters. Callers may pass the script's
# bound parameters via `Repair-OfflineDisk @PSBoundParameters`.
#
# What it does (summary):
#  - Initializes the target disk or resolves a VM to a disk
#  - Brings the disk online and assigns temporary letters as needed
#  - Detects Windows and Boot partitions
#  - Executes requested repair actions (chkdsk, DISM, BCD rebuild, RDP fixes,
#    user additions, registry fixes, etc.) according to the provided switches
#  - Optionally leaves the disk online or takes it offline when finished
function Repair-OfflineDisk {
    param (
        [string]$VMName = "",
        [int]$DiskNumber = -1,
        [switch]$FixNTFS,
        [switch]$FixBoot,
        [switch]$FixBootSector,
        [switch]$FixHealth,
        [switch]$TryLGKC,
        [switch]$TryOtherBootConfig,        
        [switch]$TrySafeMode,
        [switch]$RemoveSafeModeFlag,
        [switch]$RunSFC,
        [switch]$EnableBootLog,
        [switch]$DisableStartupRepair,
        [switch]$EnableStartupRepair,
        [switch]$DisableBFE,
        [switch]$EnableBFE,
        [switch]$AddTempUser,
        [switch]$AddTempUser2,
        [switch]$ResetLocalAdminPassword,
        [switch]$SetFullMemDump,
        [switch]$DisableNLA,
        [switch]$EnableNLA,
        [switch]$EnableWinRMHTTPS,
        [switch]$CheckRDPPolicies,
        [switch]$FixRDP,
        [switch]$FixRDPCert,
        [switch]$FixRDPPermissions,
        [switch]$FixRDPAuth,
        [switch]$FixUserRights,
        [switch]$FixPendingUpdates,
        [switch]$DisableWindowsUpdate,
        [switch]$RestoreRegistryFromRegBack,
        [switch]$EnableRegBackup,
        [switch]$DisableThirdPartyDrivers,
        [switch]$EnableThirdPartyDrivers,
        [switch]$GetServicesReport,
        [switch]$IncludeServices,
        [switch]$IssuesOnly,
        [string[]]$DisableDriver = @(),
        [string[]]$EnableDriver  = @(),
        [ValidateSet('Boot','System','Automatic','Manual','Disabled')][string]$DriverStartType = 'Manual',
        [switch]$DisableCredentialGuard,
        [switch]$EnableCredentialGuard,
        [switch]$DisableAppLocker,
        [switch]$FixSanPolicy,
        [switch]$FixAzureGuestAgent,
        [switch]$InstallAzureVMAgent,
        [switch]$FixDeviceFilters,
        [switch]$KeepDefaultFilters,
        [switch]$CopyACPISettings,
        [switch]$ScanNetBindings,
        [switch]$FixNetBindings,
        [switch]$SysCheck,
        [switch]$ResetNetworkStack,
        [switch]$EnableTestSigning,
        [switch]$DisableTestSigning,
        [switch]$CheckDiskHealth,
        [switch]$CollectEventLogs,
        [switch]$LeaveDiskOnline,
        [ValidateSet('SYSTEM','SOFTWARE','COMPONENTS','SAM','SECURITY')][string[]]$LoadHive  = @(),
        [ValidateSet('SYSTEM','SOFTWARE','COMPONENTS','SAM','SECURITY')][string[]]$UnloadHive = @(),
        [string]$DriveLetter = ''
    )

    # Identify the local OS disk (C: drive) to prevent accidental modification of the Hyper-V host
    $script:LocalOsDiskNumber = (Get-Partition -DriveLetter 'C' -ErrorAction SilentlyContinue).DiskNumber

    # Show usage/disk list when called with no actionable parameters
    if ($DiskNumber -lt 0 -and [string]::IsNullOrWhiteSpace($VMName)) {
        Write-Host @"
====================================================================
  Windows Boot Troubleshooting Script - Offline Disk Repair
====================================================================

USAGE:
  .\$(Split-Path -Leaf $PSCommandPath) -DiskNumber <N> [Options]
  .\$(Split-Path -Leaf $PSCommandPath) -VMName <Name> [Options]

PARAMETERS:
  -DiskNumber <int>      Disk number visible in Disk Management
  -VMName <string>       Hyper-V VM name (auto-detects disk, stops VM first)
  -LeaveDiskOnline       Keep disk online after repairs (default: take offline)
  -LoadHive <hive[,...]> Mount one or more offline registry hives and leave them loaded for manual
                           inspection/editing (implies disk stays online). Valid: SYSTEM SOFTWARE COMPONENTS SAM SECURITY
                           Example: -LoadHive SYSTEM  -> HKLM:\BROKENSYSTEM available in regedit
  -UnloadHive <hive[,...]> Unmount previously loaded hives and take the disk offline

--- AZURE AGENT ---------------------------------------------------------------
  -FixAzureGuestAgent    Enable Azure Guest Agent and RdAgent services (registry already present)
  -InstallAzureVMAgent   Install Azure VM Agent fully offline from host files (use when agent is missing)

--- BOOT & BCD ----------------------------------------------------------------
  -DisableStartupRepair  Stop VM from looping into WinRE on failed boot
  -EnableBootLog         Enable ntbtlog.txt boot logging
  -EnableStartupRepair   Re-enable automatic startup repair / WinRE on boot failure
  -EnableTestSigning     Enable BCD test signing (allow unsigned drivers)
  -DisableTestSigning    Disable BCD test signing
  -FixBoot               Rebuild BCD from scratch
  -FixBootSector         Repair MBR/VBR boot sector (Gen1/BIOS only; bootrec)
  -RemoveSafeModeFlag    Remove Safe Mode flag
  -TryLGKC               Switch boot to Last Known Good Control Set
  -TryOtherBootConfig    Switch boot to a different HKLM ControlSet
  -TrySafeMode           Set boot to Safe Mode (minimal)

--- DIAGNOSTICS (read-only) ---------------------------------------------------
  -CheckDiskHealth       Show disk/partition/filesystem health report
  -CheckRDPPolicies      Show current RDP auth policy values
  -CollectEventLogs      Copy guest event logs and crash dumps to C:\temp on the host
  -ScanNetBindings       Report third-party network binding components (non-ms_ ComponentId)
  -SysCheck              Full offline diagnostic scan: BCD, services, device filters, RDP, networking,
                           Azure Agent, security settings, crash artefacts - with fix suggestions

--- DISK & FILESYSTEM ---------------------------------------------------------
  -FixHealth             Run DISM ScanHealth + RestoreHealth
  -FixNTFS               Run chkdsk on the Windows partition
  -FixSanPolicy          Set SAN policy to OnlineAll (fix offline disks after migration)
  -RunSFC                Run SFC in offline mode
  -SetFullMemDump        Configure full memory dump + pagefile on C:

--- DRIVERS & DEVICE FILTERS --------------------------------------------------
  -DisableDriver <name[,name,...]>  Disable one or more named services or drivers (sets Start=4)
  -EnableDriver  <name[,name,...]>  Re-enable one or more named services or drivers
  -DriverStartType <type>           Start type for -EnableDriver: Boot, System, Automatic, Manual (default), Disabled
  -DisableThirdPartyDrivers    Disable all non-Microsoft Boot/System kernel drivers
  -EnableThirdPartyDrivers     Re-enable previously disabled third-party kernel drivers
  -GetServicesReport           List drivers from the offline ControlSet grouped by start type,
                                 flagging missing binaries and non-Microsoft vendors
    -IncludeServices             (sub-option) also include Win32 services, not just kernel drivers
    -IssuesOnly                  (sub-option) show only rows that need attention: missing binary,
                                   non-Microsoft vendor, or ErrorControl Severe/Critical (>=2)
  -FixDeviceFilters      Scan and remove unsafe UpperFilters/LowerFilters from device classes (disk/net/volume)
    -KeepDefaultFilters    (sub-option) strict mode: only inbox safe-list entries kept; removes Microsoft-signed
                             non-defaults like InDskFlt that are not standard on Azure VMs
  -CopyACPISettings      Copy legacy Hyper-V ACPI Enum entries to newer device IDs (VMBus->MSFT1000,
                           Hyper_V_Gen_Counter_V1->MSFT1002); fixes missing synthetic device detection

--- NETWORKING ----------------------------------------------------------------
  -DisableBFE            Disable Base Filtering Engine service
  -EnableBFE             Re-enable Base Filtering Engine service
  -FixNetBindings        Remove orphaned third-party network binding components (missing binary; prevents NDIS init failure)
  -ResetNetworkStack     Reset TCP/IP stack, Winsock and firewall at next boot

--- RDP & REMOTE ACCESS -------------------------------------------------------
  -DisableNLA            Disable Network Level Authentication
  -EnableNLA             Enable Network Level Authentication
  -EnableWinRMHTTPS      Configure WinRM HTTPS listener via startup script
  -FixRDP                Reset RDP registry settings to defaults
  -FixRDPAuth            Set optimal RDP/NLA/NTLM auth policy for recovery
  -FixRDPCert            Recreate the self-signed RDP certificate
  -FixRDPPermissions     Reset RDP private key and certificate service permissions

--- REGISTRY ------------------------------------------------------------------
  -EnableRegBackup             Enable periodic registry backups to RegBack folder
  -RestoreRegistryFromRegBack  Restore SYSTEM/SOFTWARE hives from RegBack backup

--- SECURITY ------------------------------------------------------------------
  -DisableAppLocker            Disable AppLocker enforcement and AppIDSvc (fixes boot blocked by bad policy)
  -DisableCredentialGuard      Disable Credential Guard and LSA protection
  -EnableCredentialGuard       Re-enable Credential Guard and LSA protection

--- USERS & RIGHTS ------------------------------------------------------------
  -AddTempUser           Add a local admin via Group Policy startup script
  -AddTempUser2          Add a local admin via Setup CmdLine (domain-joined VMs)
  -FixUserRights         Reset user rights assignments to Windows defaults
  -ResetLocalAdminPassword  Reset an existing local admin account password at next boot

--- WINDOWS UPDATE ------------------------------------------------------------
  -DisableWindowsUpdate  Disable Windows Update services to stop boot loops
  -FixPendingUpdates     Remove pending Windows Update transactions

AVAILABLE DISKS:
"@
        if ($null -ne $script:LocalOsDiskNumber) {
            Write-Host "  (Disk $script:LocalOsDiskNumber excluded - hosts the local C: drive and cannot be a repair target)" -ForegroundColor DarkYellow
        }

        # Build a disk-number -> VM-name map for all non-Off VMs
        # VHD/VHDX-backed drives don't expose DiskNumber on HardDrives directly;
        # resolve via Get-VHD which populates DiskNumber when the VHD is mounted.
        # Include ALL VMs (any state) so the listing is informational for Off VMs too.
        # A running VM tag is appended with its state when not Running.
        $diskToVM = @{}
        if (Get-Command Get-VM -ErrorAction SilentlyContinue) {
            Get-VM -ErrorAction SilentlyContinue | ForEach-Object {
                $v = $_
                $stateLabel = if ($v.State -eq 'Running') { $v.Name } else { "$($v.Name) [$($v.State)]" }
                foreach ($hd in $v.HardDrives) {
                    $resolvedNum = -1
                    if ($null -ne $hd.DiskNumber -and [int]$hd.DiskNumber -ge 0) {
                        $resolvedNum = [int]$hd.DiskNumber
                    }
                    elseif ($hd.Path -match '\.vhdx?$') {
                        $vhd = Get-VHD -Path $hd.Path -ErrorAction SilentlyContinue
                        if ($vhd -and $null -ne $vhd.DiskNumber -and [int]$vhd.DiskNumber -ge 0) {
                            $resolvedNum = [int]$vhd.DiskNumber
                        }
                    }
                    if ($resolvedNum -ge 0) {
                        $diskToVM[$resolvedNum] = $stateLabel
                    }
                }
            }
        }

        Get-Disk | Sort-Object Number | Where-Object { $null -eq $script:LocalOsDiskNumber -or $_.Number -ne $script:LocalOsDiskNumber } | ForEach-Object {
            $d = $_
            $vmTag = if ($diskToVM.ContainsKey([int]$d.Number)) { " [attached to VM: $($diskToVM[[int]$d.Number])]" } else { '' }
            Write-Host "  Disk $($d.Number): $($d.PartitionStyle) | $([math]::Round($d.Size / 1GB)) GB | $($d.FriendlyName)$vmTag"
            $parts = Get-Partition -DiskNumber $d.Number -ErrorAction SilentlyContinue
            if (-not $parts) { Write-Host "    (no partitions)"; return }
            foreach ($p in $parts) {
                $sz = [math]::Round($p.Size / 1MB)
                $ltr = if ($p.DriveLetter) { "$($p.DriveLetter):" } else { '(no letter)' }
                Write-Host "    +- Partition $($p.PartitionNumber): $ltr | $($p.Type) | $($sz) MB"
            }
        }
        Write-Host ""
        return
    }

    # Validate mutual exclusivity
    if ($DiskNumber -ge 0 -and -not [string]::IsNullOrWhiteSpace($VMName)) {
        Write-Error "Specify either -DiskNumber or -VMName, not both."
        return
    }

    # Refuse to operate on the local OS disk when -DiskNumber is supplied directly
    if ($null -ne $script:LocalOsDiskNumber -and $DiskNumber -ge 0 -and $DiskNumber -eq $script:LocalOsDiskNumber) {
        Write-Error "Disk $DiskNumber hosts the local C: drive. Refusing to modify it to prevent breaking the Hyper-V host."
        return
    }

    # Initialize logging and target (resolve VM/disk, bring disk online, detect partitions)
    # Determine if any write/repair action was requested (exclude pure read-only switches)
    $readOnlySwitches = @('SysCheck','CheckDiskHealth','ScanNetBindings','CheckRDPPolicies','CollectEventLogs','ShowLastSession','GetServicesReport')
    $hasRepairAction  = $PSBoundParameters.Keys | Where-Object { $readOnlySwitches -notcontains $_ -and $_ -notin @('VMName','DiskNumber','LeaveDiskOnline','DriveLetter','LoadHive','UnloadHive') }
    if ($hasRepairAction) {
        Write-Host "  Tip: if you haven't already, a VM snapshot or disk backup before making changes is always a safe starting point." -ForegroundColor DarkGray
        Write-Host ""
    }
    Start-ActionLog "Repair-OfflineDisk start"
    if (-not (Initialize-TargetDisk -VMName $VMName -DiskNumber $DiskNumber)) {
        return
    }

    try {
        if ($FixNTFS) { FixDiskCorruption -DriveLetter $DriveLetter }
        if ($FixBoot) { RebuildBCD }
        if ($FixBootSector) { FixBootSector }
        if ($FixHealth) { RunDismHealth }
        if ($RunSFC) { RunSFC }
        if ($TryLGKC) { SetLKGC }
        if ($TryOtherBootConfig) { RevertLKGC }
        if ($TrySafeMode) { SetSafeMode }
        if ($RemoveSafeModeFlag) { RemoveSafeMode }
        if ($EnableBootLog) { SetBootLog }
        if ($DisableStartupRepair) { DisableStartupRepair }
        if ($EnableStartupRepair) { EnableStartupRepair }
        if ($DisableBFE) { Disable-ServiceOrDriver -ServiceName BFE }
        if ($EnableBFE) { Enable-ServiceOrDriver -ServiceName BFE -StartValue 2 }
        if ($AddTempUser) {
            Write-Host "Please provide the credential (no .\ or domain\user, just user): " -ForegroundColor Yellow
            AddNewUser -WinDrive $script:WinDriveLetter -Credential (Get-Credential)
        }
        if ($AddTempUser2) {
            Write-Host "Please provide the credential (no .\ or domain\user, just user): " -ForegroundColor Yellow
            AddNewUser2 -Credential (Get-Credential)
        }
        if ($ResetLocalAdminPassword) { ResetLocalAdminPassword }
        if ($SetFullMemDump) { ConfigureFullMemDump }
        if ($FixRDP) { ResetRDPSettings }
        if ($FixRDPCert) { RecreateRDPCertificate }
        if ($FixRDPPermissions) { ResetRDPPrivKeyPermissions }
        if ($DisableNLA) { SetNLADisabled }
        if ($EnableNLA) { SetNLAEnabled }
        if ($EnableWinRMHTTPS) { SetWinRMHTTPSEnabled }
        if ($FixPendingUpdates) { ClearPendingUpdates }
        if ($DisableWindowsUpdate) { DisableWindowsUpdate }
        if ($FixUserRights) { ResetUserRights }
        if ($RestoreRegistryFromRegBack) { RestoreRegistryFromRegBack }
        if ($EnableRegBackup) { EnableRegBackup }
        if ($DisableThirdPartyDrivers) { DisableThirdPartyDrivers }
        if ($EnableThirdPartyDrivers) { EnableThirdPartyDrivers }
        if ($IncludeServices -and -not $GetServicesReport) {
            Write-Warning "-IncludeServices has no effect without -GetServicesReport."
        }
        if ($IssuesOnly -and -not $GetServicesReport) {
            Write-Warning "-IssuesOnly has no effect without -GetServicesReport."
        }
        if ($GetServicesReport) { GetServicesReport -IncludeServices:$IncludeServices -IssuesOnly:$IssuesOnly }
        foreach ($d in $DisableDriver) { if (-not [string]::IsNullOrWhiteSpace($d)) { Disable-ServiceOrDriver -ServiceName $d.Trim() } }
        if ($EnableDriver.Count -gt 0) {
            $startMap = @{ Boot=0; System=1; Automatic=2; Manual=3; Disabled=4 }
            $startInt = $startMap[$DriverStartType]
            foreach ($d in $EnableDriver) { if (-not [string]::IsNullOrWhiteSpace($d)) { Enable-ServiceOrDriver -ServiceName $d.Trim() -StartValue $startInt } }
        }
        if ($DisableCredentialGuard) { DisableCredentialGuard }
        if ($EnableCredentialGuard) { EnableCredentialGuard }
        if ($DisableAppLocker) { DisableAppLocker }
        if ($FixSanPolicy) { FixSanPolicy }
        if ($FixAzureGuestAgent) { FixAzureGuestAgent }
        if ($InstallAzureVMAgent) { InstallAzureVMAgentOffline }
        if ($KeepDefaultFilters -and -not $FixDeviceFilters) {
            Write-Warning "-KeepDefaultFilters has no effect without -FixDeviceFilters."
        }
        if ($FixDeviceFilters) { FixDeviceClassFilters -KeepDefaultFilters:$KeepDefaultFilters }
        if ($CopyACPISettings) { CopyACPISettings }
        if ($ScanNetBindings) { ScanNetAdapterBindings }
        if ($FixNetBindings) { RemoveOrphanedNetBindings }
        if ($SysCheck) { RunSystemCheck }
        if ($ResetNetworkStack) { ResetNetworkStack }
        if ($EnableTestSigning) { SetTestSigning -Enable $true }
        if ($DisableTestSigning) { SetTestSigning -Enable $false }
        if ($CheckDiskHealth) { CheckDiskHealth }
        if ($CollectEventLogs) { CollectEventLogs }
        if ($CheckRDPPolicies) { GetRdpAuthPolicySnapshot }
        if ($FixRDPAuth) { SetRdpAuthPolicyOptimal }

        # -LoadHive: mount requested hives and leave them loaded for manual inspection
        foreach ($hive in $LoadHive) {
            $OfflineWindowsPath = Join-Path $script:WinDriveLetter 'Windows'
            Write-Host "Loading offline $hive hive..." -ForegroundColor Cyan
            MountOffHive -WinPath $OfflineWindowsPath -Hive $hive
            Write-Host "  [OK] HKLM:\BROKEN$hive is now loaded. Use regedit or reg.exe to inspect/edit." -ForegroundColor Green
            Write-Host "  [!]  Run: .\RepairVM.ps1 -DiskNumber $($script:DiskNumber) -UnloadHive $hive  to unload and take the disk offline when done." -ForegroundColor Yellow
        }

        # -UnloadHive: unmount requested hives (before the disk-offline step in finally)
        foreach ($hive in $UnloadHive) {
            Write-Host "Unloading offline $hive hive..." -ForegroundColor Cyan
            UnmountOffHive -Hive $hive
            Write-Host "  [OK] HKLM:\BROKEN$hive unloaded." -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Repair-OfflineDisk encountered an error: $_"
        throw
    }
    finally {
        # When -LoadHive was used, keep the disk online so the loaded hives remain accessible
        $keepOnlineForHive = $LoadHive.Count -gt 0 -and $UnloadHive.Count -eq 0
        if (-not $LeaveDiskOnline -and -not $keepOnlineForHive -and $script:DiskNumber -ge 0) {
            Write-Host "`r`nSetting disk $script:DiskNumber offline..."
            try {
                Set-Disk -Number $script:DiskNumber -IsOffline $true
                Write-Host "[OK] Disk $script:DiskNumber is now offline and can be safely reconnected or used with the VM." -ForegroundColor Green
            }
            catch {
                Write-Warning "Could not set disk $script:DiskNumber offline: $_"
            }
        }
    }
}

# Show log entries and exit when -ShowLastSession is requested
if ($ShowLastSession) {
    Get-LastRepairSession -SessionId $SessionId -Detailed:$Detailed -All:$All -ExportTo $ExportTo
    return
}

# Invoke consolidated helper with script-bound parameters
Repair-OfflineDisk @PSBoundParameters