function Add-WordTable {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Table,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Array] $DataTable,
        [Xceed.Document.NET.AutoFit] $AutoFit,
        [Xceed.Document.NET.TableDesign] $Design,
        [Xceed.Document.NET.Direction] $Direction,
        [switch] $BreakPageAfterTable,
        [switch] $BreakPageBeforeTable,
        [nullable[int]] $MaximumColumns,
        #[string[]]$Titles = @('Name', 'Value'),
        [string] $OverwriteTitle,
        [switch] $DoNotAddTitle,
        [Xceed.Document.NET.Alignment] $TitleAlignment = [Xceed.Document.NET.Alignment]::center,
        [alias ("ColummnWidth")][float[]] $ColumnWidth = @(),
        [nullable[float]] $TableWidth = $null,
        [bool] $Percentage,
        [alias ("C")] [System.Drawing.KnownColor[]]$Color = @(),
        [alias ("S")] [double[]] $FontSize = @(),
        [alias ("FontName")] [string[]] $FontFamily = @(),
        [alias ("B")] [nullable[bool][]] $Bold = @(),
        [alias ("I")] [nullable[bool][]] $Italic = @(),
        [alias ("U")] [Xceed.Document.NET.UnderlineStyle[]] $UnderlineStyle = @(),
        [alias ('UC')] [System.Drawing.KnownColor[]]$UnderlineColor = @(),
        [alias ("SA")] [double[]] $SpacingAfter = @(),
        [alias ("SB")] [double[]] $SpacingBefore = @(),
        [alias ("SP")] [double[]] $Spacing = @(),
        [alias ("H")] [Xceed.Document.NET.Highlight[]] $Highlight = @(),
        [alias ("CA")] [Xceed.Document.NET.CapsStyle[]] $CapsStyle = @(),
        [alias ("ST")] [Xceed.Document.NET.StrikeThrough[]] $StrikeThrough = @(),
        [alias ("HT")] [Xceed.Document.NET.HeadingType[]] $HeadingType = @(),
        [int[]] $PercentageScale = @(), # "Value must be one of the following: 200, 150, 100, 90, 80, 66, 50 or 33"
        [Xceed.Document.NET.Misc[]] $Misc = @(),
        [string[]] $Language = @(),
        [int[]]$Kerning = @(), # "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72"
        [nullable[bool][]]$Hidden = @(),
        [int[]]$Position = @(), #  "Value must be in the range -1585 - 1585"
        [single[]] $IndentationFirstLine = @(),
        [single[]] $IndentationHanging = @(),
        [Xceed.Document.NET.Alignment[]] $Alignment = @(),
        [Xceed.Document.NET.Direction[]] $DirectionFormatting = @(),
        [Xceed.Document.NET.ShadingType[]] $ShadingType = @(),
        [Xceed.Document.NET.Script[]] $Script = @(),
        [nullable[bool][]] $NewLine = @(),
        [switch] $KeepLinesTogether,
        [switch] $KeepWithNextParagraph,
        [switch] $ContinueFormatting,
        [alias('Rotate', 'RotateData', 'TransposeColumnsRows', 'TransposeData')][switch] $Transpose,
        [string[]] $ExcludeProperty,
        [bool] $Supress = $false,
        [string] $Splitter = ';'
    )
    Begin {
        [int] $Run = 0
        [int] $RowNr = 0
        if ($MaximumColumns -eq $null) { $MaximumColumns = 5 }
    }
    Process {
        if ($DataTable.Count -gt 0) {
            if ($Run -eq 0) {
                if ($Transpose) {
                    $DataTable = Format-TransposeTable -Object $DataTable
                }
                $Data = Format-PSTable -Object $DataTable -ExcludeProperty $ExcludeProperty -SkipTitle:$DoNotAddTitle -Splitter $Splitter
                $WorksheetHeaders = $Data[0] # Saving Header information for later use
                $NumberRows = $Data.Count
                $NumberColumns = if ($Data[0].Count -ge $MaximumColumns) { $MaximumColumns } else { $Data[0].Count }

                if ($DoNotAddTitle) {
                    ### Add Table or Add To TABLE
                    if ($null -eq $Table) {
                        $Table = New-WordTable -WordDocument $WordDocument -Paragraph $Paragraph -NrRows ($NumberRows + 1) -NrColumns $NumberColumns -Supress $false
                    } else {
                        Add-WordTableRow -Table $Table -Count ($NumberRows + 1) -Supress $True
                    }
                } else {
                    ### Add Table or Add To TABLE
                    if ($null -eq $Table) {
                        $Table = New-WordTable -WordDocument $WordDocument -Paragraph $Paragraph -NrRows $NumberRows -NrColumns $NumberColumns -Supress $false
                    } else {
                        Add-WordTableRow -Table $Table -Count $NumberRows -Supress $True
                    }
                }
                #Write-Verbose "Add-WordTable - Run: $Run NumberRows: $NumberRows NumberColumns: $NumberColumns"
            } else {
                $Data = Format-PSTable -Object $DataTable -SkipTitle -OverwriteHeaders $WorksheetHeaders -Splitter $Splitter
                $NumberRows = $Data.Count
                $NumberColumns = if ($Data[0].Count -ge $MaximumColumns) { $MaximumColumns } else { $Data[0].Count }

                ### Add Table or Add To TABLE
                if ($null -eq $Table) {
                    $Table = New-WordTable -WordDocument $WordDocument -Paragraph $Paragraph -NrRows $NumberRows -NrColumns $NumberColumns -Supress $false
                } else {
                    Add-WordTableRow -Table $Table -Count $NumberRows -Supress $True
                }
                #Write-Verbose "Add-WordTable - Run: $Run NumberRows: $NumberRows NumberColumns: $NumberColumns"
            }
            ### Add titles
            <#
     ### Prepare Number of ROWS/COLUMNS
        $pattern = 'string|bool|byte|char|decimal|double|float|int|long|sbyte|short|uint|ulong|ushort'
        $Titles = ($DataTable | Get-Member | Where-Object { $_.MemberType -like "*Property" -and $_.Definition -match $pattern }) | Select-Object Name
        $NumberColumns = if ($Titles.Count -ge $MaximumColumns) { $MaximumColumns } else { $Titles.Count }
        $NumberRows = $DataTable.Count
        Write-Verbose 'Add-WordTable - Option 4'
        Write-Verbose "Add-WordTable - Column Count $($NumberColumns) Rows Count $NumberRows "
#>
            <#
            if (-not $DoNotAddTitle) {
                Add-WordTableTitle -Title $Titles `
                    -Table $Table `
                    -MaximumColumns $MaximumColumns `
                    -Color $Color[0] `
                    -FontSize $FontSize[0] `
                    -FontFamily $FontFamily[0] `
                    -Bold $Bold[0] `
                    -Italic $Italic[0] `
                    -UnderlineStyle $UnderlineStyle[0] `
                    -UnderlineColor $UnderlineColor[0] `
                    -SpacingAfter $SpacingAfter[0] `
                    -SpacingBefore $SpacingBefore[0] `
                    -Spacing $Spacing[0] `
                    -Highlight $Highlight[0] `
                    -CapsStyle $CapsStyle[0] `
                    -StrikeThrough $StrikeThrough[0] `
                    -HeadingType $HeadingType[0] `
                    -PercentageScale $PercentageScale[0] `
                    -Misc $Misc[0] `
                    -Language $Language[0] `
                    -Kerning $Kerning[0] `
                    -Hidden $Hidden[0] `
                    -Position $Position[0] `
                    -IndentationFirstLine $IndentationFirstLine[0] `
                    -IndentationHanging $IndentationHanging[0] `
                    -Alignment $Alignment[0] `
                    -DirectionFormatting $DirectionFormatting[0] `
                    -ShadingType $ShadingType[0] `
                    -Script $Script[0] -Supress $True
            }
#>
            ### Continue formatting
            if ($ContinueFormatting -eq $true) {
                $Formatting = Set-WordContinueFormatting -Count $NumberRows `
                    -Color $Color `
                    -FontSize $FontSize `
                    -FontFamily $FontFamily `
                    -Bold $Bold `
                    -Italic $Italic `
                    -UnderlineStyle $UnderlineStyle `
                    -UnderlineColor $UnderlineColor `
                    -SpacingAfter $SpacingAfter `
                    -SpacingBefore $SpacingBefore `
                    -Spacing $Spacing `
                    -Highlight $Highlight `
                    -CapsStyle $CapsStyle `
                    -StrikeThrough $StrikeThrough `
                    -HeadingType $HeadingType `
                    -PercentageScale $PercentageScale `
                    -Misc $Misc `
                    -Language $Language `
                    -Kerning $Kerning `
                    -Hidden $Hidden `
                    -Position $Position `
                    -IndentationFirstLine $IndentationFirstLine `
                    -IndentationHanging $IndentationHanging `
                    -Alignment $Alignment `
                    -DirectionFormatting $DirectionFormatting `
                    -ShadingType $ShadingType `
                    -Script $Script

                $Color = $Formatting[0]
                $FontSize = $Formatting[1]
                $FontFamily = $Formatting[2]
                $Bold = $Formatting[3]
                $Italic = $Formatting[4]
                $UnderlineStyle = $Formatting[5]
                $UnderlineColor = $Formatting[6]
                $SpacingAfter = $Formatting[7]
                $SpacingBefore = $Formatting[8]
                $Spacing = $Formatting[9]
                $Highlight = $Formatting[10]
                $CapsStyle = $Formatting[11]
                $StrikeThrough = $Formatting[12]
                $HeadingType = $Formatting[13]
                $PercentageScale = $Formatting[14]
                $Misc = $Formatting[15]
                $Language = $Formatting[16]
                $Kerning = $Formatting[17]
                $Hidden = $Formatting[18]
                $Position = $Formatting[19]
                $IndentationFirstLine = $Formatting[20]
                $IndentationHanging = $Formatting[21]
                $Alignment = $Formatting[22]
                $DirectionFormatting = $Formatting[23]
                $ShadingType = $Formatting[24]
                $Script = $Formatting[25]
            }
            ###  Build data in Table

            # $RowNr = 0
            #Write-Color "[i] Presenting table after conversion" -Color Yellow
            if ($Run -eq 0 -and $DoNotAddTitle) {
                #     if ($RowNr -eq 0) {
                #$Data = $Data | Select-Object -Skip 1
                #      }
                $RowNr = 1
            }
            foreach ($Row in $Data) {
                $ColumnNr = 0
                foreach ($Column in $Row) {
                    # Write-Verbose "Row: $RowNr Column: $ColumnNr Data: $Column"

                    Add-WordTableCellValue -Table $Table -Row $RowNr -Column $ColumnNr -Value $Column `
                        -Color $Color[$RowNr] `
                        -FontSize $FontSize[$RowNr] `
                        -FontFamily $FontFamily[$RowNr] `
                        -Bold $Bold[$RowNr] `
                        -Italic $Italic[$RowNr] `
                        -UnderlineStyle $UnderlineStyle[$RowNr]`
                        -UnderlineColor $UnderlineColor[$RowNr]`
                        -SpacingAfter $SpacingAfter[$RowNr] `
                        -SpacingBefore $SpacingBefore[$RowNr] `
                        -Spacing $Spacing[$RowNr] `
                        -Highlight $Highlight[$RowNr] `
                        -CapsStyle $CapsStyle[$RowNr] `
                        -StrikeThrough $StrikeThrough[$RowNr] `
                        -HeadingType $HeadingType[$RowNr] `
                        -PercentageScale $PercentageScale[$RowNr] `
                        -Misc $Misc[$RowNr] `
                        -Language $Language[$RowNr]`
                        -Kerning $Kerning[$RowNr]`
                        -Hidden $Hidden[$RowNr]`
                        -Position $Position[$RowNr]`
                        -IndentationFirstLine $IndentationFirstLine[$RowNr]`
                        -IndentationHanging $IndentationHanging[$RowNr]`
                        -Alignment $Alignment[$RowNr]`
                        -DirectionFormatting $DirectionFormatting[$RowNr] `
                        -ShadingType $ShadingType[$RowNr]`
                        -Script $Script[$RowNr] -Supress $True
                    if ($ColumnNr -eq $($MaximumColumns - 1)) {
                        # prevents display of more columns then there is space, choose carefully
                        break;
                    }
                    $ColumnNr++

                }
                #if ($DataTable.Count -eq $RowNr) {
                $RowNr++
            }
            $Run++
        }
    }
    End {
        if ($DataTable.Count -gt 0) {
            ### Apply formatting to table

            $Table | Set-WordTableColumnWidth -Width $ColumnWidth -TotalWidth $TableWidth -Percentage $Percentage -Supress $True

            $Table | Set-WordTable -Direction $Direction `
                -AutoFit $AutoFit `
                -Design $Design `
                -BreakPageAfterTable:$BreakPageAfterTable `
                -BreakPageBeforeTable:$BreakPageBeforeTable -Supress $True

            if ($OverwriteTitle) {
                $Table = Set-WordTableRowMergeCells -Table $Table -RowNr 0 -MergeAll
                $TableParagraph = Get-WordTableRow -Table $Table -RowNr 0 -ColumnNr 0
                $TableParagraph = Set-WordText -Paragraph $TableParagraph -Text $OverwriteTitle -Alignment $TitleAlignment
            }

            ### return data
            if ($Supress) { return } else { return $Table }
        }
    }
}

