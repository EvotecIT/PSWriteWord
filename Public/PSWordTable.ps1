function Add-WordTableTitle {
    [CmdletBinding()]
    param(
        $Table,
        $Titles,
        $MaximumColumns
    )
    Write-Verbose "Add-WordTableTitle - Title Count $($Titles.Count) "

    #$Titles

    #Write-Color "Title Count $($Titles.Count) " -Color Yellow
    for ($a = 0; $a -lt $Titles.Count; $a++) {
        if ($Titles[$a] -is [string]) {
            #$Titles[$a].GetType()
            $ColumnName = $Titles[$a]
        } else {
            $ColumnName = $Titles[$a].Name
        }
        Write-Verbose "Add-WordTableTitle - Column Name: $ColumnName"
        Add-WordTableCellValue -Table $Table -Row 0 -Column $a -Value $ColumnName -Supress $Supress
        if ($a -eq $($MaximumColumns - 1)) {
            break;
        }
    }
}
function Add-WordTableCellValue {
    [CmdletBinding()]
    param(
        $Table,
        $Row,
        $Column,
        $Value,
        $Paragraph = 0,
        [bool] $Supress = $true
    )
    Write-Verbose "Add-WordTableCellValue - Row: $Row Column $Column Value $Value"
    $Data = $Table.Rows[$Row].Cells[$Column].Paragraphs[$Paragraph].Append($Value)
    if ($Supress -eq $true) { return } else { return $Data }
}
function Add-WordTable {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)] [Xceed.Words.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [ValidateNotNullOrEmpty()]$Table,
        [TableDesign] $Design = [TableDesign]::ColorfulList,
        [int] $MaximumColumns = 5,
        [string[]]$Columns = @('Name', 'Value'),
        [bool] $Supress = $true
    )

    if ($Table.GetType().BaseType.Name -eq 'Array' -and $Table.GetType().Name -eq 'Object[]') {
        Write-Verbose 'Add-WordTable - Converting Array of Objects'
        $Table = $Table.ForEach( {[PSCustomObject]$_})
    }
    $ObjectType = $Table.GetType().Name
    Write-Verbose "Add-WordTable - Table row count: $(Get-ObjectCount $table)"
    Write-Verbose "Add-WordTable - Object Type: $ObjectType"
    Write-Verbose "Add-WordTable - BaseType.Name: $($Table.GetType().BaseType.Name)"
    Write-Verbose "Add-WordTable - GetType Before Conversion: $ObjectType"

    If ($ObjectType -eq 'Hashtable' -or $ObjectType -eq 'OrderedDictionary') {

    } else {
        $Table = $Table | Select-Object *
    }

    $ObjectType = $Table.GetType().Name

    Write-Verbose "Add-WordTable - GetType After Conversion: $ObjectType"

    if ($ObjectType -eq 'Hashtable' -or $ObjectType -eq 'OrderedDictionary') {
        Write-Verbose 'Add-WordTable - Option 1'
        $NumberRows = $Table.Count + 1
        $NumberColumns = 2

        Write-Verbose "Add-WordTable - Column Count $($NumberColumns) Rows Count $NumberRows "
        Write-Verbose "Add-WordTable - Titles: $([string] $Columns)"

        if ($Paragraph -eq $null) {
            $WordTable = $WordDocument.InsertTable($NumberRows, $NumberColumns)
        } else {
            $TableDefinition = $WordDocument.AddTable($NumberRows, $NumberColumns)
            $WordTable = $Paragraph.InsertTableAfterSelf($TableDefinition)
        }

        Add-WordTableTitle -Title $Columns -Table $WordTable -MaximumColumns $MaximumColumns
        $Row = 1
        foreach ($TableEntry in $Table.GetEnumerator()) {
            $ColumnNrForTitle = 0
            $ColumnNrForData = 1
            $Data = Add-WordTableCellValue -Table $WordTable -Row $Row -Column $ColumnNrForTitle -Value $TableEntry.Name
            $Data = Add-WordTableCellValue -Table $WordTable -Row $Row -Column $ColumnNrForData -Value $TableEntry.Value
            Write-Verbose "Add-WordTable - RowNr: $Row / ColumnNr: $ColumnTitle Name: $($TableEntry.Name) Value: $($TableEntry.Value)"
            $Row++

        }
    } elseif ($ObjectType -eq 'PSCustomObject') {
        Write-Verbose 'Add-WordTable - Option 2'

        $Titles = Get-ObjectTitles -Object $Table[0]

        $NumberRows = $Titles.Count + 1
        $NumberColumns = 2

        Write-Verbose "Add-WordTable - Column Count $($NumberColumns) Rows Count $NumberRows "
        Write-Verbose "Add-WordTable - Titles: $([string] $Titles)"

        $WordTable = New-WordTable -WordDocument $WordDocument -Paragraph $Paragraph -NrRows $NumberRows -NrColumns $NumberColumns -Supress $false

        Add-WordTableTitle -Title $Columns -Table $WordTable -MaximumColumns $MaximumColumns
        $Row = 1
        foreach ($Title in $Titles) {
            $Value = Get-ObjectData -Object $Table -Title $Title -DoNotAddTitles

            $ColumnTitle = 0
            $ColumnData = 1
            $Data = Add-WordTableCellValue -Table $WordTable -Row $Row -Column $ColumnTitle -Value $Title
            $Data = Add-WordTableCellValue -Table $WordTable -Row $Row -Column $ColumnData -Value $Value
            Write-Verbose "Add-WordTable - Title:  $Title Value: $Value Row: $Row "
            $Row++

        }
    } elseif ($Table.GetType().Name -eq 'Object[]') {
        write-verbose 'Add-WordTable - option 3'

        $Titles = Get-ObjectTitles -Object $Table[0]

        $NumberColumns = if ($Titles.Count -ge $MaximumColumns) { $MaximumColumns } else { $Titles.Count }
        $NumberRows = $Table.Count + 1

        Write-Verbose "Add-WordTable - Column Count $($NumberColumns) Rows Count $NumberRows "
        Write-Verbose "Add-WordTable - Titles: $([string] $Titles)"
        #Write-Color "Column Count ", $NumberColumns, " Rows Count ", $NumberRows -C Yellow, Green, Yellow, Green

        $WordTable = New-WordTable -WordDocument $WordDocument -Paragraph $Paragraph -NrRows $NumberRows -NrColumns $NumberColumns -Supress $false

        Add-WordTableTitle -Title $Titles -Table $WordTable -MaximumColumns $MaximumColumns

        for ($b = 0; $b -lt $NumberRows - 1; $b++) {
            $a = 0
            foreach ($Title in $Titles) {
                $Data = Add-WordTableCellValue -Table $WordTable -Row $($b + 1) -Column $a -Value $Table[$b].$Title
                if ($a -eq $($MaximumColumns - 1)) { break; } # prevents display of more columns then there is space, choose carefully
                $a++
            }
        }
    } else {
        Write-Verbose 'Add-WordTable - Option 4'
        $pattern = 'string|bool|byte|char|decimal|double|float|int|long|sbyte|short|uint|ulong|ushort'
        $Columns = ($Table | Get-Member | Where-Object { $_.MemberType -like "*Property" -and $_.Definition -match $pattern }) | Select-Object Name
        #$Columns
        $NumberColumns = if ($Columns.Count -ge $MaximumColumns) { $MaximumColumns } else { $Columns.Count }
        $NumberRows = $Table.Count

        Write-Verbose "Add-WordTable - Column Count $($NumberColumns) Rows Count $NumberRows "
        #Write-Color "Column Count ", $NumberColumns, " Rows Count ", $NumberRows -C Yellow, Green, Yellow, Green

        $WordTable = New-WordTable -WordDocument $WordDocument -Paragraph $Paragraph -NrRows $NumberRows -NrColumns $NumberColumns -Supress $false

        $Titles = Add-WordTableTitle -Title $Columns -Table $WordTable -MaximumColumns $MaximumColumns

        for ($b = 1; $b -lt $NumberRows; $b++) {
            $a = 0
            foreach ($Title in $Columns.Name) {
                $Data = Add-WordTableCellValue -Table $WordTable -Row $b -Column $a -Value $Table[$b].$Title
                if ($a -eq $($MaximumColumns - 1)) { break; } # prevents display of more columns then there is space, choose carefully
                $a++


            }
        }

    }

    $WordTable.Design = $Design


    if ($Supress -eq $false) { return $WordTable } else { return }
}

function New-WordTableBorder {
    [CmdletBinding()]
    param (
        [BorderStyle] $BorderStyle,
        [BorderSize] $BorderSize,
        [int] $BorderSpace,
        [System.Drawing.Color] $BorderColor
    )

    $Border = New-Object -TypeName Xceed.Words.NET.Border -ArgumentList $BorderStyle, $BorderSize, $BorderSpace, $BorderColor
    return $Border
}

function Set-WordTableBorder {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [TableBorderType] $TableBorderType,
        [Border] $Border
    )
    $Table.SetBorder($TableBorderType, $Border)
}

function New-WordTable {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)] [Xceed.Words.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [int] $NrRows,
        [int] $NrColumns,
        [bool] $Supress = $true
    )

    if ($Paragraph -eq $null) {
        $WordTable = $WordDocument.InsertTable($NrRows, $NrColumns)
    } else {
        $TableDefinition = $WordDocument.AddTable($NrRows, $NrColumns)
        $WordTable = $Paragraph.InsertTableAfterSelf($TableDefinition)
    }
    if ($Supress) { return } else { return $WordTable }
}