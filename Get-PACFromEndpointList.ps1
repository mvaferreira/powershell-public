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
#>

Param (
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [String] $DirectProxySettings = 'DIRECT',

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [String] $DefaultProxySettings = 'PROXY 10.10.10.10:8080',

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [String] $DefaultProxyO365Settings = 'PROXY 10.20.20.20:8080'    
)

$guid = [GUID]::NewGuid().Guid
$EndpointsURL = "https://endpoints.office.com/endpoints/worldwide?clientrequestid=" + $guid + "&format=CSV&NoIPv6=true"
$directProxyVarName = "direct"
$defaultProxyVarName = "proxyServer"
$proxyO365 = "proxyO365"
$bl = "`r`n"
$tab = "`t"
$InternalFile = ".\proxyDirect.txt"
$BlackListFile = ".\blacklist.txt"

If (Test-Path -Path $BlackListFile) {
    $BlackList = Get-Content -Path $BlackListFile
}

Function ConvertTo-DottedDecimalIP {
    <#
      .Synopsis
        Returns a dotted decimal IP address from either an unsigned 32-bit integer or a dotted binary string.
      .Description
        ConvertTo-DottedDecimalIP uses a regular expression match on the input string to convert to an IP address.
      .Parameter IPAddress
        A string representation of an IP address from either UInt32 or dotted binary.
    #>
  
    [CmdLetBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [String]$IPAddress
    )
    
    process {
        Switch -RegEx ($IPAddress) {
            "([01]{8}.){3}[01]{8}" {
                return [String]::Join('.', $( $IPAddress.Split('.') | ForEach-Object { [Convert]::ToUInt32($_, 2) } ))
            }
            "\d" {
                $IPAddress = [UInt32]$IPAddress
                $DottedIP = $( For ($i = 3; $i -gt -1; $i--) {
                        $Remainder = $IPAddress % [Math]::Pow(256, $i)
                        ($IPAddress - $Remainder) / [Math]::Pow(256, $i)
                        $IPAddress = $Remainder
                    } )
         
                return [String]::Join('.', $DottedIP)
            }
            default {
                Write-Error "Cannot convert this format"
            }
        }
    }
}

Function ConvertTo-Mask {
    <#
      .Synopsis
        Returns a dotted decimal subnet mask from a mask length.
      .Description
        ConvertTo-Mask returns a subnet mask in dotted decimal format from an integer value ranging 
        between 0 and 32. ConvertTo-Mask first creates a binary string from the length, converts 
        that to an unsigned 32-bit integer then calls ConvertTo-DottedDecimalIP to complete the operation.
      .Parameter MaskLength
        The number of bits which must be masked.
    #>
    
    [CmdLetBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [Alias("Length")]
        [ValidateRange(0, 32)]
        $MaskLength
    )
    
    Process {
        return ConvertTo-DottedDecimalIP ([Convert]::ToUInt32($(("1" * $MaskLength).PadRight(32, "0")), 2))
    }
}

Function Get-PacClauses {
    param(
        [Parameter(Mandatory = $false)]
        [string[]] $Urls,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ReturnVarName
    )

    if (!$Urls) {
        return ""
    }

    $clauses = (($Urls | ForEach-Object {
                If ($_) {
                    If ($_ -notin $BlackList) {
                        "shExpMatch(host, `"$_`")"
                    }
                }
            }) -Join "$bl$tab|| ")

    @"
    if($clauses)
    {
        return $ReturnVarName;
    }
"@
}

Function Get-PacIPClauses {
    param(
        [Parameter(Mandatory = $false)]
        [string[]] $IPs,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ReturnVarName
    )

    if (!$IPs) {
        return ""
    }

    $clauses = (($IPs | ForEach-Object {
                If ($_) {
                    $IPMask = $_ -split "/"
                    $IP = $IPMask[0]
                    $MaskBits = [int]$IPMask[1]

                    If ($IP -notin $BlackList) {
                        "isInNet(host, `"$IP`", `"$(ConvertTo-Mask $MaskBits)`")"
                    }
                }
            }) -Join "$bl$tab|| ")

    @"
    if($clauses)
    {
        return $ReturnVarName;
    }
"@
}

Function Get-PacString {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [array[]] $MapVarUrls,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [array[]] $MapVarIPs,
        
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [array[]] $MapVarIntUrls,
        
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [array[]] $MapVarIntIPs 
    )

    @"
// This PAC file will provide proxy config to Microsoft 365 services
// using data from the public web service for all endpoints
function FindProxyForURL(url, host)
{
    var $directProxyVarName = "$DirectProxySettings";
    var $defaultProxyVarName = "$DefaultProxySettings";
    var $proxyO365 = "$DefaultProxyO365Settings";
 
$( (Get-PACClauses -ReturnVarName $proxyO365 -Urls $MapVarUrls) -Join "$bl$bl" )

$( (Get-PacIPClauses -ReturnVarName $proxyO365 -IPs $MapVarIPs) -Join "$bl$bl" )

$( (Get-PacClauses -ReturnVarName $directProxyVarName -Urls $MapVarIntUrls) -Join "$bl$bl" )

$( (Get-PacIPClauses -ReturnVarName $directProxyVarName -IPs $MapVarIntIPs) -Join "$bl$bl" )
 
    return $defaultProxyVarName;
}
"@ -replace "($bl){3,}", "$bl$bl" # Collapse more than one blank line in the PAC file so it looks better.
}

$varUrls = ""
$varIPs = ""
$varIntUrls = ""
$varIntIps = ""

Try {
    $CSVContent = (ConvertFrom-Csv (Invoke-WebRequest -Uri $EndpointsURL).Content)

    If ($CSVContent) {
        ForEach ($data In $CSVContent) {
            Switch -Regex ($data.category) {
                "Allow|Optimize|Default" {
                    If ($data.urls) {
                        $varUrls += $data.urls + ","
                    }

                    If ($data.ips) {
                        $varIPs += $data.ips + ","
                    }
                }
            }
        }
    }

    If (Test-Path $InternalFile) {
        $InternalAddresses = Get-Content -Path $InternalFile

        If ($InternalAddresses) {
            ForEach($Addr In $InternalAddresses) {
                If (-Not $Addr.Contains("/")) {
                    $varIntUrls += $Addr + ","
                } Else {
                    $varIntIps += $Addr + ","
                }
            }
        }
    }
}
Catch {
    $Error[0].Exception
}

Get-PacString -MapVarUrls $($varUrls -split "," | Sort-Object | Get-Unique) `
                -MapVarIPs $($varIPs -split "," | Sort-Object | Get-Unique) `
                -MapVarIntUrls $($varIntUrls -split "," | Sort-Object | Get-Unique) `
                -MapVarIntIPs $($varIntIps -split "," | Sort-Object | Get-Unique)