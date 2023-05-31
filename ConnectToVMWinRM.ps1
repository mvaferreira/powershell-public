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
    This script will enable WinRM over HTTPS port 5986 through Azure RunCommand feature.

    .EXAMPLES
    #Enable WinRM over HTTPS and connect using private IP address
    .\ConnectToVMWinRM.ps1 -VMName <vmname> -RGName <rgname>

    #WinRM over HTTPS is already enabled, just try to connect
    .\ConnectToVMWinRM.ps1 -VMName <vmname> -RGName <rgname> -ConnectOnly

    #Collect credential to reuse and attempt to connect
    $Cred = Get-Credential
    .\ConnectToVMWinRM.ps1 -VMName <vmname> -RGName <rgname> -Credential $Cred

    #Try to enable WinRM and attempt to connect using public IP address (assuming NSG allows inbound tcp/5986)
    .\ConnectToVMWinRM.ps1 -VMName <vmname> -RGName <rgname> -UsePublicIP
#>

Param (
    [Parameter(Mandatory = $True)]
    [string]$VMName = "",

    [Parameter(Mandatory = $True)]
    [string]$RGName = "",

    [Parameter(Mandatory = $False)]
    [switch]$UsePublicIP = $False,

    [Parameter(Mandatory = $False)]
    [switch]$ConnectOnly = $False,    

    [Parameter(Mandatory = $False)]
    [pscredential]$Credential = ""   
)

$VM = Get-AzVM -Name $VMName -ResourceGroupName $RGName -ErrorAction SilentlyContinue
$VMStatus = Get-AzVM -Name $VMName -ResourceGroupName $RGName -Status -ErrorAction SilentlyContinue
$VMPowerState = ($VMStatus.Statuses | Where-Object { $_.Code -match 'Powerstate' }).Code
$AgentStatus = $VMStatus.VMAgent.Statuses.DisplayStatus

If ($VM -And ($VMPowerState -Match "running") -And ($AgentStatus -Match "Ready")) {
    If ($UsePublicIP) {
        $PubIPId = (Get-AzNetworkInterface -ResourceId $VM.NetworkProfile.NetworkInterfaces.Id).IpConfigurations.PublicIpAddress.Id
        $PIPAddress = (Get-AzPublicIpAddress | Where-Object { $_.Id -eq $PubIPId }).IpAddress

        If (-Not $PIPAddress) {
            Write-Host "Could not find a public IP in VM $($VMName)."
            Return
        }
    }
    Else {
        $IPAddress = (Get-AzNetworkInterface -ResourceId $VM.NetworkProfile.NetworkInterfaces.Id).IpConfigurations[0].PrivateIpAddress
    }

    If (-Not $Credential) {
        $Credential = Get-Credential
    }

    If (-Not $ConnectOnly) {
        $EnableWinRM = {
            Enable-PSRemoting -Force
            Enable-WSManCredSSP -Role "Server" -Force

            If (-Not (Get-NetFirewallRule -DisplayName "Windows Remote Management (HTTPS-In)" -ErrorAction SilentlyContinue)) {
                $FirewallParam = @{
                    DisplayName = 'Windows Remote Management (HTTPS-In)'
                    Direction   = 'Inbound'
                    LocalPort   = 5986
                    Protocol    = 'TCP'
                    Action      = 'Allow'
                    Program     = 'System'
                }
                New-NetFirewallRule @FirewallParam | Out-Null
            }

            $Cert = New-SelfSignedCertificate -DnsName (hostname) -CertStoreLocation Cert:\LocalMachine\My

            If (-Not (Get-ChildItem -Recurse WSMan:\localhost\Listener | Where-Object { $_.Value -match "HTTPS" -Or $_.Value -match "5986" })) {
                New-Item WSMan:\localhost\Listener -Address * -Transport HTTPS -HostName (hostname) -CertificateThumbPrint $Cert.Thumbprint -Port 5986 -Force | Out-Null
            }
        }

        New-Item -Path . -Name "EnableWinRM.ps1" -ItemType File -Value $EnableWinRM -Force | Out-Null
        Write-Host "[$(Get-Date)] Enabling WinRM over HTTPS..."
        Invoke-AzVMRunCommand -ResourceGroupName $RGName -VMName $VMName -CommandId "RunPowerShellScript" -ScriptPath ".\EnableWinRM.ps1" | Out-Null
        If (Test-Path -Path ".\EnableWinRM.ps1") { Remove-Item -Path ".\EnableWinRM.ps1" -Force | Out-Null }
    }

    Write-Host "[$(Get-Date)] Trying to connect to VM $($VMName)..."

    If ($UsePublicIP) {
        Enter-PSSession -ConnectionUri "https://$($PIPAddress):5986" -Credential $Credential -SessionOption (New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck)
    }
    Else {
        Enter-PSSession -ConnectionUri "https://$($IPAddress):5986" -Credential $Credential -SessionOption (New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck)
    }
}
Else {
    Write-Host "Verify that VM $($VMName) exist in RG $($RGName), is started and agent status is ready."
}