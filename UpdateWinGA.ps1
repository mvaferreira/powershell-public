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

Get-ChildItem -Directory -Filter "*GuestAgent*" -Path $WinGAFolder | ForEach {
    Move-Item -Path $_.FullName -Destination "$($WinGAFolder)\OLD" -Force
}

Get-ChildItem -Directory -Filter "Packages*" -Path $WinGAFolder | ForEach {
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
    Catch { }
}