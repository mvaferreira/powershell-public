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
    Version: 0.2

    .DESCRIPTION
    This script will nicely format the OnDemandAssessment (ODA) excel report.
    Please refer to the following url for more information about OnDemandAssessments.
    https://docs.microsoft.com/en-us/services-hub/health/getting_started_with_on_demand_assessments
    
    .EXAMPLE
    Provide the report path to the -Report parameter.
    .\FormatODAReport.ps1 -Report "C:\temp\AssessmentPlanReport_WindowsServerAssessment.xlsx"
#>

Param(
    [string] $Report = ""
)

#Return if no report available to process
If ($Report -And ($Report.Length -gt 1)) {
    If (-Not (Test-Path -Path $Report)) {
        Write-Host "File $($Report) not found."
        return
    }
}
Else {
    Write-Host "Specify a report to process."
    return
}

#Get excel letter sequence for column index
Function ExcelSeq($col) {
    While ($col -gt 0) {
        $curLetterNum = ($col - 1) % 26;
        $curLetter = [char]$([int]$curLetterNum + 65)
        $colString = $curLetter + $colString
        $col = ($col - ($curLetterNum + 1)) / 26
    }
    return $colString
}

Function Get-RGB { 
    Param( 
        [Parameter(Mandatory = $false)] 
        [ValidateRange(0, 255)] 
        [Int] 
        $Red = 0, 
        [Parameter(Mandatory = $false)] 
        [ValidateRange(0, 255)] 
        [Int] 
        $Green = 0,
        [Parameter(Mandatory = $false)] 
        [ValidateRange(0, 255)] 
        [Int] 
        $Blue = 0 
    ) 
    Process { 
        [long]($Red + ($Green * 256) + ($Blue * 65536))
    } 
}

Function CreatePivotTable($destSheet, $DTName, $pivotName, $WSObj, $RowFields, $DataField, $DTFieldSummary, $DTFieldText, $ColumnFields) {
    $xlDatabase = 1
    $xlPivotTableVersion12 = 3
    $PivotTable = $WB.PivotCaches().Create($xlDatabase, $DTName, $xlPivotTableVersion12)
    Start-Sleep -Milliseconds 500
    $destName = "'" + $destSheet + "'!R1C1"
    $PivotTable.CreatePivotTable([string]$destName, $pivotName) | out-null
    
    $I = 1
    ForEach ($Field In $RowFields) {
        $WSObj.PivotTables($pivotName).PivotFields($Field).Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
        $WSObj.PivotTables($pivotName).PivotFields($Field).Position = $I
        $I++
    }

    $null = $WSObj.PivotTables($pivotName).AddDataField($WSObj.PivotTables($pivotName).PivotFields([string]$DataField),
        [string]($DTFieldText), $DTFieldSummary)

    #We have to loop again, after row fields were added
    $I = 1
    ForEach ($Field In $RowFields) {
        If ($I -ne $RowFields.Count) {
            $WSObj.PivotTables($pivotName).PivotFields($Field).ShowDetail = $False
        }
        $I++
    }

    If ($ColumnFields) {
        $I = 1
        ForEach ($Field In $ColumnFields) {
            $WSObj.PivotTables($pivotName).PivotFields($Field).Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlColumnField
            $WSObj.PivotTables($pivotName).PivotFields($Field).Position = $I
            $I++
        }

        $I = 1
        ForEach ($Field In $ColumnFields) {
            If ($I -ne $ColumnFields.Count) {
                $WSObj.PivotTables($pivotName).PivotFields($Field).ShowDetail = $False
            }
            $I++
        }
    }

    [int]$ChartWidth = 600
    [int]$ChartHeight = 400
    $chartType = [Microsoft.Office.Interop.Excel.XlChartType]::xlColumnStacked

    $chart = $WSObj.Shapes.AddChart().Chart
    $chart.ShowReportFilterFieldButtons = $False
    $chart.ShowLegendFieldButtons = $False
    $chart.ShowAxisFieldButtons = $False
    $chart.ShowValueFieldButtons = $False
    $chart.ShowAllFieldButtons = $False
    $chart.ShowExpandCollapseEntireFieldButtons = $False
    $chart.ChartType = $chartType
    $chart.HasTitle = $True
    $chart.ChartTitle.Text = $destSheet

    $chart.SeriesCollection(1) | ForEach-Object {
        If ($_.XValues) {
            $I = 1
            ForEach ($Value In $_.XValues) {
                Switch ($Value) {
                    'Catastrophic' {
                        $chart.SeriesCollection(1).Points($I).Format.Fill.ForeColor.RGB = $(Get-RGB 139 0 0)
                    }

                    'Very High' {
                        $chart.SeriesCollection(1).Points($I).Format.Fill.ForeColor.RGB = $(Get-RGB 178 34 34)
                    }
                    
                    'High' {
                        $chart.SeriesCollection(1).Points($I).Format.Fill.ForeColor.RGB = $(Get-RGB 255 0 0)
                    }

                    'Moderate' {
                        $chart.SeriesCollection(1).Points($I).Format.Fill.ForeColor.RGB = $(Get-RGB 255 140 0)
                    }

                    'Low to Moderate' {
                        $chart.SeriesCollection(1).Points($I).Format.Fill.ForeColor.RGB = $(Get-RGB 218 165 32)
                    }

                    'Low' {
                        $chart.SeriesCollection(1).Points($I).Format.Fill.ForeColor.RGB = $(Get-RGB 30 144 255)
                    }

                    'Very Low' {
                        $chart.SeriesCollection(1).Points($I).Format.Fill.ForeColor.RGB = $(Get-RGB 135 206 250)
                    }                                     
                }
                $I++
            }
        }
    }

    $WSObj.Shapes.Item("Chart 1").Top = 50
    $WSObj.Shapes.Item("Chart 1").Width = $ChartWidth
    $WSObj.Shapes.Item("Chart 1").Height = $ChartHeight
    
}

#Logic Main
Try {
    Write-Host -NoNewline "Formatting worksheet..."
    
    #Save new excel.exe PID
    $AllPIDs = Get-Process excel -ErrorAction Ignore | ForEach-Object { $_.Id }
    $XL = New-Object -ComObject Excel.Application
    $ExcelPID = Get-Process excel -ErrorAction Ignore | ForEach-Object { $_.Id } | Where-Object { $AllPIDs -notcontains $_ }

    $XL.Visible = $False
    $WB = $XL.Workbooks.Open($Report)
    $DataSheetName = "AssessmentWorkSheet"
    $WS = $WB.Worksheets.Item($DataSheetName)

    #Get used row and column count
    $RowCount = $WS.UsedRange.Rows.Count
    $ColumnCount = $WS.UsedRange.Columns.Count

    #UsedArea
    #Starts with cell A3, skipping sheet title and headers
    #UsedArea = A3:L(numOfRows) (Ex. 'A3:L142')
    $UsedArea = [string]$(ExcelSeq(1)) + "3:" + [string]$(ExcelSeq($ColumnCount)) + $RowCount

    #Merge title cells and apply alignment
    $WS.Range("A1:" + [string]$(ExcelSeq($ColumnCount)) + "1").MergeCells = $True
    $WS.Range("A1:" + [string]$(ExcelSeq($ColumnCount)) + "1").HorizontalAlignment = -4131
    $WS.Range("A1:" + [string]$(ExcelSeq($ColumnCount)) + "1").VerticalAlignment = -4160

    $DataTableName = "TableAssessmentFindings"

    #If table already exist, delete it
    If ($WB.ActiveSheet.ListObjects.Count -gt 0) {
        If ($WB.ActiveSheet.ListObjects($DataTableName)) {
            $WB.ActiveSheet.ListObjects($DataTableName).Unlist()
        }
    }

    #Apply table style to used area
    $ListObject = $WB.ActiveSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,
        $WS.Range("A2:" + [string]$(ExcelSeq($ColumnCount)) + $RowCount),
        $null , [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
    $ListObject.Name = $DataTableName
    $ListObject.TableStyle = "TableStyleMedium6"

    #Set fixed row height to 180
    For ($Row = 3; $Row -le $RowCount; $Row++) {
        $WS.Cells.EntireRow.Item($Row).RowHeight = 180
    }

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
                $WS.Cells.Item(2, $Column).HorizontalAlignment = -4131

                For ($Row = 3; $Row -le $RowCount; $Row++) {
                    Switch ($WS.Cells.Item($Row, $Column).Value2) {
                        'Failed' {
                            $WS.Cells.Item($Row, $Column).Style = "Bad"
                        }

                        'Passed' {
                            $WS.Cells.Item($Row, $Column).Style = "Good"
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
                #Set column format as number
                $WS.Cells.EntireColumn.Item($Column).NumberFormat = "0.0"
                $WS.Cells.EntireColumn.Item($Column).ColumnWidth = 16

                For ($Row = 3; $Row -le $RowCount; $Row++) {
                    $CurValue = [string]$WS.Cells.Item($Row, $Column).Value2
                    $WS.Cells.Item($Row, $Column).Value2 = $CurValue
                }

                #Sort column descending
                $ColumnLetter = [string]$(ExcelSeq($Column))
                $null = $WS.Range($UsedArea).Sort($WS.Range($($ColumnLetter + 3)),
                    [Microsoft.Office.Interop.Excel.XlSortOrder]::xlDescending,
                    $WS.Range($($ColumnLetter + $RowCount)),
                    $null,
                    [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlNo,
                    $null,
                    [Microsoft.Office.Interop.Excel.XlSortDataOption]::xlSortTextAsNumbers,
                    [Microsoft.Office.Interop.Excel.XlSortDataOption]::xlSortTextAsNumbers)              

                $WS.Cells.EntireColumn.Item($Column).HorizontalAlignment = -4108
                $WS.Cells.EntireColumn.Item($Column).VerticalAlignment = -4160
                $WS.Cells.Item(2, $Column).HorizontalAlignment = -4131
            }

            'Probability' {
                $WS.Cells.EntireColumn.Item($Column).ColumnWidth = 16
                $WS.Cells.EntireColumn.Item($Column).HorizontalAlignment = -4108
                $WS.Cells.EntireColumn.Item($Column).VerticalAlignment = -4160
                $WS.Cells.Item(2, $Column).HorizontalAlignment = -4131
            }

            'Impact' {
                $WS.Cells.EntireColumn.Item($Column).ColumnWidth = 16
                $WS.Cells.EntireColumn.Item($Column).HorizontalAlignment = -4108
                $WS.Cells.EntireColumn.Item($Column).VerticalAlignment = -4160
                $WS.Cells.Item(2, $Column).HorizontalAlignment = -4131

                For ($Row = 3; $Row -le $RowCount; $Row++) {
                    Switch ($WS.Cells.Item($Row, $Column).Value2) {
                        'Catastrophic' {
                            # BackgroundColor = RGB 139 0 0
                            # ForegroundColor = Black (1)
                            $WS.Cells.Item($Row, $Column).Interior.Color = $(Get-RGB 139 0 0)
                            $WS.Cells.Item($Row, $Column).Font.ColorIndex = 1
                        }

                        'Very High' {
                            # BackgroundColor = RGB 178 34 34
                            # ForegroundColor = Black (1)
                            $WS.Cells.Item($Row, $Column).Interior.Color = $(Get-RGB 178 34 34)
                            $WS.Cells.Item($Row, $Column).Font.ColorIndex = 1
                        }

                        'High' {
                            # BackgroundColor = RGB 255 0 0
                            # ForegroundColor = Black (1)
                            $WS.Cells.Item($Row, $Column).Interior.Color = $(Get-RGB 255 0 0)
                            $WS.Cells.Item($Row, $Column).Font.ColorIndex = 1
                        }

                        'Moderate' {
                            # BackgroundColor = RGB 255 140 0
                            # ForegroundColor = Black (1)
                            $WS.Cells.Item($Row, $Column).Interior.Color = $(Get-RGB 255 140 0)
                            $WS.Cells.Item($Row, $Column).Font.ColorIndex = 1
                        }

                        'Low to Moderate' {
                            # BackgroundColor = RGB 218 165 32
                            # ForegroundColor = Black (1)
                            $WS.Cells.Item($Row, $Column).Interior.Color = $(Get-RGB 218 165 32)
                            $WS.Cells.Item($Row, $Column).Font.ColorIndex = 1
                        }

                        'Low' {
                            # BackgroundColor = RBG 30 144 255
                            # ForegroundColor = Black (1)
                            $WS.Cells.Item($Row, $Column).Interior.Color = $(Get-RGB 30 144 255)
                            $WS.Cells.Item($Row, $Column).Font.ColorIndex = 1
                        }

                        'Very Low' {
                            # BackgroundColor = RBG 135 206 250
                            # ForegroundColor = Black (1)
                            $WS.Cells.Item($Row, $Column).Interior.Color = $(Get-RGB 135 206 250)
                            $WS.Cells.Item($Row, $Column).Font.ColorIndex = 1
                        }
                    }
                    
                }
            }            

            'Effort' {
                $WS.Cells.EntireColumn.Item($Column).ColumnWidth = 16
                $WS.Cells.EntireColumn.Item($Column).HorizontalAlignment = -4108
                $WS.Cells.EntireColumn.Item($Column).VerticalAlignment = -4160
                $WS.Cells.Item(2, $Column).HorizontalAlignment = -4131
            }

            'Technology' {
                $WS.Cells.EntireColumn.Item($Column).ColumnWidth = 16
                $WS.Cells.EntireColumn.Item($Column).WrapText = $True
                $WS.Cells.EntireColumn.Item($Column).HorizontalAlignment = -4108
                $WS.Cells.EntireColumn.Item($Column).VerticalAlignment = -4160
                $WS.Cells.Item(2, $Column).HorizontalAlignment = -4131
            }

            'Source' {
                $WS.Cells.EntireColumn.Item($Column).ColumnWidth = 16
                $WS.Cells.EntireColumn.Item($Column).HorizontalAlignment = -4108
                $WS.Cells.EntireColumn.Item($Column).VerticalAlignment = -4160
                $WS.Cells.Item(2, $Column).HorizontalAlignment = -4131
            }
        }
    }

    #Add PivotTables, if only found the Assessment worksheet

    If ($WB.Worksheets.Count -eq 2) {
        #region 'Recommendations by Effort'
        $WSName = "Recommendations by Effort"
        $WSPT = $WB.Worksheets.Add()
        $WSPT.Name = $WSName

        $pivotTableName = "EffortPivotTable"
        $RowFields = @("Effort", "Focus Area", "Content and Best Practices")
        $DataField = "Score"
        $DataFieldText = [string]"Count of Effort"
        $DataFieldSummary = [Microsoft.Office.Interop.Excel.XlConsolidationFunction]::xlCount
        CreatePivotTable $WSName $DataTableName $pivotTableName $WSPT $RowFields $DataField $DataFieldSummary $DataFieldText
        #EndRegion 'Recommendations by Effort'

        #region 'Recommendations by Impact'
        $WSName = "Recommendations by Impact"
        $WSPT = $WB.Worksheets.Add()
        $WSPT.Name = $WSName

        $pivotTableName = "ImpactPivotTable"
        $RowFields = @("Impact", "Recommendation Title", "Content and Best Practices")
        $DataField = "Score"
        $DataFieldText = [string]"Count of Impact"
        $DataFieldSummary = [Microsoft.Office.Interop.Excel.XlConsolidationFunction]::xlCount
        CreatePivotTable $WSName $DataTableName $pivotTableName $WSPT $RowFields $DataField $DataFieldSummary $DataFieldText
        #EndRegion 'Recommendations by Impact'

        #region 'Recommendations by Focus Area'
        $WSName = "Recommendations by Focus Area"
        $WSPT = $WB.Worksheets.Add()
        $WSPT.Name = $WSName

        $pivotTableName = "FocusAreaPivotTable"
        $RowFields = @("Focus Area", "Recommendation Title", "Content and Best Practices", "Affected Objects")
        $DataField = "Score"
        $DataFieldText = [string]"Count of Focus Area"
        $DataFieldSummary = [Microsoft.Office.Interop.Excel.XlConsolidationFunction]::xlCount
        CreatePivotTable $WSName $DataTableName $pivotTableName $WSPT $RowFields $DataField $DataFieldSummary $DataFieldText
        #EndRegion 'Recommendations by Focus Area'

        $LastSheet = $WB.Worksheets | Select-Object -Last 1
        $LastSheet.Move($WSPT)
    }
    Else {
        Write-Host "`r`nWarning! Make sure this report is not already formatted and there is only one worksheet.`r`nSkipping pivot tables worksheets."
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

    Write-Host "done!"
    Write-Host "New version available at $($NewFileName)"
}
Catch {
    Write-Host "`r`n"
    Write-Host $Error[0].Exception.Message

    #Make sure COM variables are released, so excel.exe is gone.
    If ($WS) { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($WS) }
    If ($WB) { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($WB) }
    If ($XL) { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($XL) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
Finally {
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

#Hold while excel.exe is still running. Try to kill it.
Do {
    Stop-Process -Id $ExcelPID -Force -ErrorAction Ignore
    Start-Sleep -Milliseconds 100
} While (Get-Process -Id $ExcelPID -ErrorAction Ignore)