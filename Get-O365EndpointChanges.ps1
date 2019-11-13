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
    This script will get all changes to Microsoft Office 365 endpoint list, since specified version.

    REFERENCES
    * Doc URL: https://docs.microsoft.com/en-us/office365/enterprise/urls-and-ip-address-ranges
    * RSS URL: https://endpoints.office.com/version/worldwide?allversions=true&format=rss&clientrequestid=<GUID>
    * Changes URL: https://endpoints.office.com/changes/worldwide/0000000000?ClientRequestId=<GUID>
    
    .EXAMPLE
    Get the latest changes
    .\Get-O365Changes.ps1

    .EXAMPLE
    Get the latest changes, since version "0000000000"
    .\Get-O365Changes.ps1 -version "0000000000"

    .EXAMPLE
    Get the latest changes, since version "0000000000", with action "Add"
    .\Get-O365Changes.ps1 -version "0000000000" -action "Add"
#>

Param (
    $version = "",
    $action = ""
)

If ($action) {
    If ("Add", "Change", "Remove" -notcontains $action) {
        Write-Host "Specify action, either 'Add', 'Change' or 'Remove'"
        break
    }
}

$AllChanges = @()

Try {
    $guid = [GUID]::NewGuid().Guid
    $RSSURL = "https://endpoints.office.com/version/worldwide?allversions=true&format=rss&clientrequestid=" + $guid

    $RSSContent = (Invoke-WebRequest -Uri $RSSURL).Content

    #Get current and previous endpoint change version
    If ($RSSContent) {
        $xmlContent = [xml]$RSSContent

        If ($version) {
            $previousVersion = $version
        }
        Else {
            $previousVersion = $xmlContent.rss.channel.item.guid.'#text' | Sort-Object -Descending | Select-Object -Skip 1 | Select-Object -First 1
        }

        Write-Host "Previous Version: " $previousVersion

        #Get changes from previous to current version as CSV        
        If ($previousVersion) {
            $changesURL = "https://endpoints.office.com/changes/worldwide/" + $previousVersion + "?ClientRequestId=" + $guid + "&format=CSV"
            $csvContent = (ConvertFrom-Csv (Invoke-WebRequest -Uri $changesURL).Content)

            Write-Host "Current Version: " $csvContent[$csvContent.Length - 1].version

            #Split changes into Add, Change and Remove items            
            If ($csvContent) {
                Foreach ($data In $csvContent) {                  
                    Switch ($data.disposition) {
                        'Add' {
                            $AllChanges += [PSCustomObject]@{
                                action          = $data.disposition;
                                impact          = $data.impact;
                                currentTcpPorts = $data.currentTcpPorts;
                                currentUdpPorts = $data.currentUdpPorts;
                                urlsAdded       = $data.urlsAdded;
                                ipsAdded        = $data.ipsAdded;
                                version         = $data.version;
                            }                            
                        }

                        'Change' {
                            $AllChanges += [PSCustomObject]@{
                                action           = $data.disposition;
                                impact           = $data.impact;
                                previousTcpPorts = $data.previousTcpPorts;
                                currentTcpPorts  = $data.currentTcpPorts;
                                previousUdpPorts = $data.previousUdpPorts;
                                currentUdpPorts  = $data.currentUdpPorts;
                                urlsAdded        = $data.urlsAdded;
                                ipsAdded         = $data.ipsAdded;
                                urlsRemoved      = $data.urlsRemoved;
                                ipsRemoved       = $data.ipsRemoved;
                                version          = $data.version;
                            }
                        }

                        'Remove' {
                            $AllChanges += [PSCustomObject]@{
                                action           = $data.disposition;
                                impact           = $data.impact;
                                previousTcpPorts = $data.previousTcpPorts;
                                currentTcpPorts  = $data.currentTcpPorts;
                                previousUdpPorts = $data.previousUdpPorts;
                                currentUdpPorts  = $data.currentUdpPorts;
                                urlsRemoved      = $data.urlsRemoved;
                                ipsRemoved       = $data.ipsRemoved;
                                version          = $data.version;
                            }
                        }
                    }
                }
            }
        }
    }
}
Catch {
    $Error[0].Exception
}

#Either display all changes or requested ones.
If ($action) {
    $AllChanges | Where-Object { $_.action -eq $action }
}
Else {
    $AllChanges
}