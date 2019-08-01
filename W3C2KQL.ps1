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
    Author: Marcus Ferreira <marcus.ferreira@microsoft.com>
    Version: 0.1

    .DESCRIPTION
    This script will parse a w3c log file to extract its header and convert into Kusto query Language (KQL).
    It helps getting started with custom logs in Azure Monitor, Log Analytics.
    
    .EXAMPLE
    Specify a w3c log file with a header to the -File parameter.
    .\W3C2KQL.ps1 -File "C:\temp\mylog.w3c"

    Go to your Log Analytics workspace, add a custom log, specify a w3c log file.

    Warning: make sure your w3c log is a tsv (tab separated value),
    otherwise '\t' won't be able to split the header.

    Run the KQL as:

    MyCustomData
    | extend fields = split(RawData, '\t')
    | extend ['computer'] = tostring(fields[0])
    | extend ['datetime'] = todatetime(strcat(fields[1], ' ', fields[2]))
    | project ['computer'],['datetime']
#>

Param(
    [string] $File = ""
)

If (-Not $File) {
    Write-Host "Specify a w3c log file using -File parameter."
    return
}

$FirstFiveLines = Get-Content -Path $File -Head 5

If ($FirstFiveLines | Select-String "username") {
    If ($FirstFiveLines | Select-String "#Fields") {
        $tHeader = ($FirstFiveLines | Select-String "#Fields").toString().Trim()
        $Header = $tHeader.Substring(9,$tHeader.Length-9)
    }
    else {
        $Header = ($FirstFiveLines | Select-String "username").toString().Trim()
    }

    $i=-1
    $t = $null
    Write-Host "| extend fields = split(RawData, '\t')"
    $Header.Split("`t") | ForEach-Object {
        $i++

        If ($_.ToString().ToLower() -ne "time") {
            If ($_.ToString().ToLower() -ne "date") {
                Write-Host "`| extend `[`'$($_)`'`] = tostring(fields[$($i)])"
                If ($t) {
                    $t = $t + ",`[`'$($_)`'`]"
                } else {
                    $t = $t + "`[`'$($_)`'`]"
                }
            } Else {
                Write-Host "`| extend `[`'datetime'`] = todatetime(strcat(fields`[$($i)`], ' ', fields`[$($i+1)`]))"
                $t = $t + ",`[`'datetime'`]"
            }
        }
    }

    Write-Host "`| project $($t)"
} Else {
    Write-Host "Header not found. Ensure the file contains a header."
    return
}