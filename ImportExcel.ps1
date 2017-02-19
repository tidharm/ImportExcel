function ConvertFrom-ExcelColumnName {
    param($columnName)

    $sum=0
    $columnName.ToCharArray() |
        ForEach-Object {
            $sum*=26
            $sum+=[char]$_.tostring().toupper()-[char]'A'+1
        } 
    $sum
}

function ConvertTo-RowCol {
    param($Range)

    if($Range -match "(?'StartCol'[a-zA-z])(?'StartRow'\d+):(?'EndCol'[a-zA-z])(?'EndRow'\d+)") {
        $StartCol=ConvertFrom-ExcelColumnName $Matches.StartCol
        $StartRow=$Matches.StartRow
        $EndCol=ConvertFrom-ExcelColumnName $Matches.EndCol
        $EndRow=$Matches.EndRow
        [PSCustomObject][Ordered]@{
            StartCol=$StartCol
            StartRow=$StartRow
            EndCol=$EndCol
            EndRow=$EndRow
            Rows=$EndRow-$StartRow+1
            Columns=$EndCol-$StartCol+1
        }
    }
}

function Import-Excel {
    <#
    .SYNOPSIS
        Read the content of an Excel sheet.
 
    .DESCRIPTION 
        The Import-Excel cmdlet reads the content of an Excel worksheet and creates one object for each row. This is done without using Microsoft Excel in the background but by using the .NET EPPLus.dll. You can also automate the creation of Pivot Tables and Charts.
 
    .PARAMETER Path 
        Specifies the path to the Excel file.
 
    .PARAMETER WorkSheetname
        Specifies the name of the worksheet in the Excel workbook. 
        
    .PARAMETER HeaderRow
        Specifies custom header names for columns.

    .PARAMETER Header
        Specifies the title used in the worksheet. The title is placed on the first line of the worksheet.

    .PARAMETER NoHeader
        When used we generate our own headers (P1, P2, P3, ..) instead of the ones defined in the first row of the Excel worksheet.

    .PARAMETER DataOnly
        When used we will only generate objects for rows that contain text values, not for empty rows or columns.
 
    .EXAMPLE
        Import-Excel -WorkSheetname 'Statistics' -Path 'E:\Finance\Company results.xlsx'
        Imports all the information found in the worksheet 'Statistics' of the Excel file 'Company results.xlsx'

    .LINK
        https://github.com/dfinke/ImportExcel
    #>
    param(
        [Alias("FullName")]
        [Parameter(ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, Mandatory=$true)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        $Path,
        [Alias("Sheet")]
        $WorkSheetname=1,
        [OfficeOpenXml.ExcelAddress]$Range,
        [int]$HeaderRow=1,
        [string[]]$Header,
        [switch]$NoHeader,
        [switch]$DataOnly
    )

    Process {

        $Path = (Resolve-Path $Path).ProviderPath
        write-debug "target excel file $Path"

        $stream = New-Object -TypeName System.IO.FileStream -ArgumentList $Path,"Open","Read","ReadWrite"
        $xl = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $stream

        $workbook  = $xl.Workbook

        $worksheet=$workbook.Worksheets[$WorkSheetname]
        $dimension=$worksheet.Dimension

        if($Range) {
            $TargetRowsCols=ConvertTo-RowCol $Range

            $HeaderRow=$TargetRowsCols.StartRow
            $StartCol=$TargetRowsCols.StartCol-1
            #$Rows=$TargetRowsCols.Rows+1
            $Rows=$TargetRowsCols.Rows+$HeaderRow
            $Columns=$TargetRowsCols.Columns+$TargetRowsCols.StartCol-1
        } else {
            $Rows=$dimension.Rows
            $Columns=$dimension.Columns
            $StartCol=0
        }
    
        if ($NoHeader) {
            if ($DataOnly) {
                $CellsWithValues = $worksheet.Cells | Where-Object Value

                $Script:i = 0
                $ColumnReference = $CellsWithValues | Select-Object -ExpandProperty End | Group-Object Column |
                Select-Object @{L='Column';E={$_.Name}}, @{L='NewColumn';E={$Script:i++; $Script:i}}
                
                $CellsWithValues | Select-Object -ExpandProperty End | Group-Object Row | ForEach-Object {    
                    $newRow = [Ordered]@{}
                    
                    foreach ($C in $ColumnReference) {
                        $newRow."P$($C.NewColumn)" = $worksheet.Cells[($_.Name),($C.Column)].Value
                    }

                    [PSCustomObject]$newRow
                }
            }
            else {
                foreach ($Row in $HeaderRow..($Rows-1)) {
                    $newRow = [Ordered]@{}
                    foreach ($Column in $StartCol..($Columns-1)) {
                        $propertyName = "P$($Column+1)"
                        $newRow.$propertyName = $worksheet.Cells[($Row+1),($Column+1)].Value
                    }

                    [PSCustomObject]$newRow
                }
            }
        } 
        else {
            if (!$Header) {
                $Header = foreach ($Column in 1..$Columns) {
                    $worksheet.Cells[$HeaderRow,$Column].Value
                }
            }

            if ($Rows -eq 1) {
                $Header | ForEach-Object {$h=[Ordered]@{}} {$h.$_=''} {[PSCustomObject]$h}
            } 
            else {
                if ($DataOnly) {
                    $CellsWithValues = $worksheet.Cells | Where-Object {$_.Value -and ($_.End.Row -ne 1)}

                    $Script:i = -1
                    $ColumnReference = $CellsWithValues | Select-Object -ExpandProperty End | Group-Object Column |
                    Select-Object @{L='Column';E={$_.Name}}, @{L='NewColumn';E={$Script:i++; $Header[$Script:i]}}
                
                    $CellsWithValues | Select-Object -ExpandProperty End | Group-Object Row | ForEach-Object {    
                        $newRow = [Ordered]@{}
                    
                        foreach ($C in $ColumnReference) {
                            $newRow."$($C.NewColumn)" = $worksheet.Cells[($_.Name),($C.Column)].Value
                        }

                        [PSCustomObject]$newRow
                    }
                }
                else {
                    foreach ($Row in ($HeaderRow+1)..$Rows) {
                        $h=[Ordered]@{}
                        foreach ($Column in $StartCol..($Columns-1)) {
                            if($Header[$Column].Length -gt 0) {
                                $Name    = $Header[$Column]
                                $h.$Name = $worksheet.Cells[$Row,($Column+1)].Value
                            }
                        }
                        [PSCustomObject]$h
                    }
                }
            }
        }
        

        $stream.Close()
        $stream.Dispose()
        $xl.Dispose()
        $xl = $null
    }
}