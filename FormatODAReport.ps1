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
    This script will nicely format the OnDemandAssessment (ODA) excel report.
    
    .EXAMPLE
    Provide the report path to the -Report parameter.
    .\FormatODAReport.ps1 -Report "C:\temp\AssessmentPlanReport_WindowsServerAssessment.xlsx"
#>

Param(
    [string] $Report = ""
)

If ($Report -And ($Report.Length -gt 1)) {
    If (-Not (Test-Path -Path $Report)) {
        Write-Host "File $($Report) not found."
        return
    }
} Else {
    Write-Host "Specify a report to process."
    return
}

Function ExcelSeq($col) {
    While ($col -gt 0) {
        $curLetterNum = ($col - 1) % 26;
        $curLetter = [char]$([int]$curLetterNum + 65)
        $colString = $curLetter + $colString
        $col = ($col - ($curLetterNum + 1)) / 26
    }
    return $colString
}

Try {
    #Save new excel.exe PID
    $AllPIDs = Get-Process excel -ErrorAction Ignore | ForEach-Object { $_.Id }
    $XL = New-Object -ComObject Excel.Application
    $ExcelPID = Get-Process excel -ErrorAction Ignore | ForEach-Object { $_.Id } | Where-Object { $AllPIDs -notcontains $_ }

    $XL.Visible = $False
    $WB = $XL.Workbooks.Open($Report)
    $WS = $WB.Worksheets.Item("AssessmentWorkSheet")

    $RowCount = $WS.UsedRange.Rows.Count
    $ColumnCount = $WS.UsedRange.Columns.Count

    #UsedArea
    #Starts with cell A3, skipping sheet title and headers
    #UsedArea = A3:L(numOfRows) (Ex. 'A3:L142')
    $UsedArea = [string]$(ExcelSeq(1)) + "3:" + [string]$(ExcelSeq($ColumnCount)) + $RowCount

    #Logic Main
    Write-Host -NoNewline "Formatting worksheet..."

    #Merge title cells and apply alignment
    $WS.Range("A1:" + [string]$(ExcelSeq($ColumnCount)) + "1").MergeCells = $True
    $WS.Range("A1:" + [string]$(ExcelSeq($ColumnCount)) + "1").HorizontalAlignment = -4131
    $WS.Range("A1:" + [string]$(ExcelSeq($ColumnCount)) + "1").VerticalAlignment = -4160

    #Apply table style to used area
    $ListObject = $WB.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $WS.Range("A2:" + [string]$(ExcelSeq($ColumnCount)) + $RowCount), $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
    $ListObject.Name = "TableData"
    $ListObject.TableStyle = "TableStyleMedium6"

    #Loop through all columns and format them as needed
    For ($Column = 1; $Column -le $ColumnCount; $Column++) {
        $WS.Cells.EntireColumn.Item($Column).HorizontalAlignment = -4131
        $WS.Cells.EntireColumn.Item($Column).VerticalAlignment = -4160

        Switch ($WS.Cells.Item(2, $Column).Value2) {
            'Recommendation Title' {
                $WS.Cells.EntireColumn.Item($Column).ColumnWidth = 40
                $WS.Cells.EntireColumn.Item($Column).WrapText = $True
            }

            'Why Consider This' {
                $WS.Cells.EntireColumn.Item($Column).ColumnWidth = 60
                $WS.Cells.EntireColumn.Item($Column).WrapText = $True
            }

            'Focus Area' {
                $WS.Cells.EntireColumn.Item($Column).ColumnWidth = 25
            }

            'Status' {
                $WS.Cells.EntireColumn.Item($Column).ColumnWidth = 16
                $WS.Cells.EntireColumn.Item($Column).HorizontalAlignment = -4108
                $WS.Cells.EntireColumn.Item($Column).VerticalAlignment = -4160
                $WS.Cells.Item(2,$Column).HorizontalAlignment = -4131

                For ($Row = 3; $Row -le $RowCount; $Row++) {
                    Switch ($WS.Cells.Item($Row, $Column).Value2) {
                        'Failed' {
                            $WS.Cells.Item($Row, $Column).Interior.ColorIndex = 15
                            $WS.Cells.Item($Row, $Column).Font.ColorIndex = 3
                        }

                        'Passed' {
                            $WS.Cells.Item($Row, $Column).Interior.ColorIndex = 4
                            $WS.Cells.Item($Row, $Column).Font.ColorIndex = 1
                        }
                    }
                }
            }

            'Content and Best Practices' {
                $WS.Cells.EntireColumn.Item($Column).ColumnWidth = 60
                $WS.Cells.EntireColumn.Item($Column).WrapText = $True
            }

            'Affected Objects' {
                $WS.Cells.EntireColumn.Item($Column).WrapText = $True
                $WS.Cells.EntireColumn.Item($Column).ColumnWidth = 35
            }

            'Score' {
                $ColumnLetter = [string]$(ExcelSeq($Column))
                $null = $WS.Range($UsedArea).Sort($WS.Range($($ColumnLetter + 3)), 2, $WS.Range($($ColumnLetter + $RowCount)), $null, 1, $null, 2, 2)
                $WS.Cells.EntireColumn.Item($Column).NumberFormat = "0.00"
                $WS.Cells.EntireColumn.Item($Column).ColumnWidth = 16

                For ($Row = 3; $Row -le $RowCount; $Row++) {
                    $CurValue = [string]$WS.Cells.Item($Row, $Column).Value2
                    $WS.Cells.Item($Row, $Column).Value2 = $CurValue
                }

                $WS.Cells.EntireColumn.Item($Column).HorizontalAlignment = -4108
                $WS.Cells.EntireColumn.Item($Column).VerticalAlignment = -4160
                $WS.Cells.Item(2,$Column).HorizontalAlignment = -4131
            }

            'Probability' {
                $WS.Cells.EntireColumn.Item($Column).ColumnWidth = 16
                $WS.Cells.EntireColumn.Item($Column).HorizontalAlignment = -4108
                $WS.Cells.EntireColumn.Item($Column).VerticalAlignment = -4160
                $WS.Cells.Item(2,$Column).HorizontalAlignment = -4131
            }

            'Impact' {
                $WS.Cells.EntireColumn.Item($Column).ColumnWidth = 16
                $WS.Cells.EntireColumn.Item($Column).HorizontalAlignment = -4108
                $WS.Cells.EntireColumn.Item($Column).VerticalAlignment = -4160
                $WS.Cells.Item(2,$Column).HorizontalAlignment = -4131

                For ($Row = 3; $Row -le $RowCount; $Row++) {
                    Switch ($WS.Cells.Item($Row, $Column).Value2) {
                        'Catastrophic' {
                            # BackgroundColor = Red (3)
                            # ForegroundColor = Black (1)
                            $WS.Cells.Item($Row, $Column).Interior.ColorIndex = 3
                            $WS.Cells.Item($Row, $Column).Font.ColorIndex = 1
                        }

                        'Very High' {
                            # BackgroundColor = Red (3)
                            # ForegroundColor = Black (1)
                            $WS.Cells.Item($Row, $Column).Interior.ColorIndex = 3
                            $WS.Cells.Item($Row, $Column).Font.ColorIndex = 1
                        }

                        'High' {
                            # BackgroundColor = Dark Orange (46)
                            # ForegroundColor = Black (1)
                            $WS.Cells.Item($Row, $Column).Interior.ColorIndex = 46
                            $WS.Cells.Item($Row, $Column).Font.ColorIndex = 1
                        }

                        'Moderate' {
                            # BackgroundColor = Light Orange (45)
                            # ForegroundColor = Black (1)
                            $WS.Cells.Item($Row, $Column).Interior.ColorIndex = 45
                            $WS.Cells.Item($Row, $Column).Font.ColorIndex = 1
                        }

                        'Low to Moderate' {
                            # BackgroundColor = Yellow (6)
                            # ForegroundColor = Black (1)
                            $WS.Cells.Item($Row, $Column).Interior.ColorIndex = 6
                            $WS.Cells.Item($Row, $Column).Font.ColorIndex = 1
                        }

                        'Low' {
                            # BackgroundColor = Green (4)
                            # ForegroundColor = Black (1)
                            $WS.Cells.Item($Row, $Column).Interior.ColorIndex = 4
                            $WS.Cells.Item($Row, $Column).Font.ColorIndex = 1
                        }

                        'Very Low' {
                            # BackgroundColor = Light Blue (8)
                            # ForegroundColor = Black (1)
                            $WS.Cells.Item($Row, $Column).Interior.ColorIndex = 8
                            $WS.Cells.Item($Row, $Column).Font.ColorIndex = 1
                        }
                    }
                    
                }
            }            

            'Effort' {
                $WS.Cells.EntireColumn.Item($Column).ColumnWidth = 16
                $WS.Cells.EntireColumn.Item($Column).HorizontalAlignment = -4108
                $WS.Cells.EntireColumn.Item($Column).VerticalAlignment = -4160
                $WS.Cells.Item(2,$Column).HorizontalAlignment = -4131
            }

            'Technology' {
                $WS.Cells.EntireColumn.Item($Column).ColumnWidth = 16
                $WS.Cells.EntireColumn.Item($Column).WrapText = $True
                $WS.Cells.EntireColumn.Item($Column).HorizontalAlignment = -4108
                $WS.Cells.EntireColumn.Item($Column).VerticalAlignment = -4160
                $WS.Cells.Item(2,$Column).HorizontalAlignment = -4131
            }

            'Source' {
                $WS.Cells.EntireColumn.Item($Column).ColumnWidth = 16
                $WS.Cells.EntireColumn.Item($Column).HorizontalAlignment = -4108
                $WS.Cells.EntireColumn.Item($Column).VerticalAlignment = -4160
                $WS.Cells.Item(2,$Column).HorizontalAlignment = -4131
            }
        }
    }

    #Save and close workbook
    $ReportObj = Get-Item $Report
    $NewFileName = Join-Path $ReportObj.Directory $ReportObj.Name.Replace(".xlsx", "_edited.xlsx")
    
    If (Test-Path -Path $NewFileName) {
        Remove-Item -Path $NewFileName -Force -ErrorAction Ignore
    }

    $WB.SaveAs($NewFileName)
    $WB.Close($True)

    #Close Excel
    $XL.Quit()

    #Make sure COM variables are released, so excel.exe is gone.
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($WS)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($WB)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($XL)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
Catch {
    Write-Host "`r`n"
    Write-Host $Error[0].Exception

    #Make sure COM variables are released, so excel.exe is gone.
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($WS)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($WB)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($XL)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

#Hold while excel.exe is still running. Try to kill it.
Do {
    Stop-Process -Id $ExcelPID -Force -ErrorAction Ignore
    Start-Sleep -Milliseconds 100
} While (Get-Process -Id $ExcelPID -ErrorAction Ignore)

Write-Host "done!"