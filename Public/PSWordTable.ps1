function Add-WordTableTitle ($Table, $Titles, $MaximumColumns) {
    #Write-Debug "Title Count $($Titles.Count) "
    #Write-Color "Title Count $($Titles.Count) " -Color Yellow
    for ($a = 0; $a -lt $Titles.Count; $a++) {
        if ($Titles[$a] -is [string]) {
            #$Titles[$a].GetType()
            $ColumnName = $Titles[$a]
        } else {
            $ColumnName = $Titles[$a].Name
        }
        #Write-Color "Column Name: $ColumnName" -Color DarkBlue
        Add-WordTableCellValue -Table $Table -Row 0 -Column $a -Value $ColumnName
        if ($a -eq $($MaximumColumns - 1)) {
            break;
        }
    }
}
function Add-WordTableCellValue ($Table, $Row, $Column, $Value, $Paragraph = 0) {
    #Write-Debug "Add-CellValue: $Row $Column $Value"
    #Write-Color "Add-CellValue: $Row $Column $Value" -Color Yellow
    $Table.Rows[$Row].Cells[$Column].Paragraphs[$Paragraph].Append($Value) | Out-Null
}
function Add-WordTable {
    param (
        $WordDocument,
        $Table,
        $Design = 'ColorfulList',
        [int] $MaximumColumns = 5
    )
    #Write-Color 'Table count: ', $Table.Count -Color White, Yellow
    #$Table.GetType()
    Write-Color "GetType1: ", $Table.GetType().Name -Color Yellow, White
    $Table = $Table | Select-Object *
    Write-Color "GetType2: ", $Table.GetType().Name -Color Yellow, White
    if ($Table.GetType().Name -eq 'PSCustomObject') {
        $Titles = Get-ObjectTitles -Object $Table

        $NumberRows = $Titles.Count + 1
        $NumberColumns = 2

        $WordTable = $WordDocument.InsertTable($NumberRows, $NumberColumns)
        $WordTable.Design = $Design

        $Columns = 'Name', 'Value'

        Add-WordTableTitle -Title $Columns -Table $WordTable -MaximumColumns $MaximumColumns
        $Row = 1
        foreach ($Title in $Titles) {
            $Value = Get-ObjectData -Object $Table -Title $Title -DoNotAddTitles

            $ColumnTitle = 0
            $ColumnData = 1
            Add-WordTableCellValue -Table $WordTable -Row $Row -Column $ColumnTitle -Value $Title
            Add-WordTableCellValue -Table $WordTable -Row $Row -Column $ColumnData -Value $Value
            $Row++
            Write-Color 'Title: ', $Title, ' Value: ', $Value, ' Row: ', $Row -Color Yellow, White, Yellow, White
        }

    } else {
        $pattern = 'string|bool|byte|char|decimal|double|float|int|long|sbyte|short|uint|ulong|ushort'
        $Columns = ($Table | Get-Member | Where-Object { $_.MemberType -like "*Property" -and $_.Definition -match $pattern }) | Select-Object Name

        $NumberColumns = if ($Columns.Count -ge $MaximumColumns) { $MaximumColumns } else { $Columns.Count }
        $NumberRows = $Table.Count

        Write-Debug "Column Count $($NumberColumns) Rows Count $NumberRows "
        Write-Color "Column Count ", $NumberColumns, " Rows Count ", $NumberRows -C Yellow, Green, Yellow, Green

        $WordTable = $WordDocument.InsertTable($NumberRows, $NumberColumns)
        $WordTable.Design = $Design

        Add-WordTableTitle -Title $Columns -Table $WordTable -MaximumColumns $MaximumColumns

        for ($b = 1; $b -lt $NumberRows; $b++) {
            $a = 0
            foreach ($Title in $Columns.Name) {
                Add-WordTableCellValue -Table $WordTable -Row $b -Column $a -Value $Table[$b].$Title
                if ($a -eq $($MaximumColumns - 1)) { break; } # prevents display of more columns then there is space, choose carefully
                $a++

            }
        }
    }
}