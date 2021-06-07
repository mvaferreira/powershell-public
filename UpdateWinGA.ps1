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
    This script will attempt to remove existing Windows Azure Guest Agent and upgrade it
    to the latest available from: https://go.microsoft.com/fwlink/?linkid=394789&clcid=0x409

    Please follow this document for manual steps:
    https://docs.microsoft.com/en-us/troubleshoot/azure/virtual-machines/windows-azure-guest-agent#step-2-check-whether-auto-update-is-working
    
    .EXAMPLE
    .\UpdateWinGA.ps1

    Check result in C:\WindowsAzure\WinGAInstall.log
#>

$WinGAFolder = "C:\WindowsAzure"
$Services = "RdAgent", "WindowsAzureGuestAgent", "WindowsAzureTelemetryService"
$Processes = "WindowsAzureGuestAgent.exe", "WaAppAgent.exe"
$WinGAURL = "https://go.microsoft.com/fwlink/?linkid=394789&clcid=0x409"
$WinGAMSI = "WindowsAzureVmAgent-new.msi"

ForEach ($Service In $Services) {
    If (Get-Service -Name $Service -ErrorAction SilentlyContinue) {
        Stop-Service -Name $Service -Force
    }
}

ForEach ($Process In $Processes) {
    If (Get-Process -Name $Process -ErrorAction SilentlyContinue) {
        Stop-Process -Name $Process -Force
    }
}

ForEach ($Service In $Services) {
    If (Get-Service -Name $Service -ErrorAction SilentlyContinue) {
        & cmd /c sc delete $Service
    }
}

If (-Not (Test-Path -Path "$($WinGAFolder)\OLD")) {
    New-Item -ItemType Directory -Name "OLD" -Path $WinGAFolder
}

Get-ChildItem -Directory -Filter "*GuestAgent*" -Path $WinGAFolder | ForEach-Object {
    Move-Item -Path $_.FullName -Destination "$($WinGAFolder)\OLD" -Force
}

Get-ChildItem -Directory -Filter "Packages*" -Path $WinGAFolder | ForEach-Object {
    Move-Item -Path $_.FullName -Destination "$($WinGAFolder)\OLD" -Force
}

If (-Not (Get-ChildItem -File -Filter "WindowsAzureVMAgent*msi" -Path $WinGAFolder)) {
    Invoke-WebRequest -Uri $WinGAURL -OutFile "$($WinGAFolder)\$WinGAMSI"
}

If (Test-Path -Path "$($WinGAFolder)\$WinGAMSI") {
    Unblock-File -Path "$($WinGAFolder)\$WinGAMSI"

    $params = @()
    $params += '/quiet'
    $params += '/L*v'
    $params += "$($WinGAFolder)\WinGAInstall.log"
    $params += '/i'
    $params += "$($WinGAFolder)\$WinGAMSI"
        
    Try {
        $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo 
        $ProcessInfo.FileName = "C:\windows\system32\msiexec.exe"
        $ProcessInfo.RedirectStandardError = $true
        $ProcessInfo.RedirectStandardOutput = $true
        $ProcessInfo.UseShellExecute = $false
        $ProcessInfo.Arguments = $params
        $Process = New-Object System.Diagnostics.Process
        $Process.StartInfo = $ProcessInfo
        $Process.Start() | Out-Null
        $Process.WaitForExit()
        $ReturnMSG = $Process.StandardOutput.ReadToEnd()
        $ReturnMSG
    }
    Catch { 
        Write-Host $Error[0].Exception.Message
    }
}