function Add-WordTable {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)] [Xceed.Words.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [Xceed.Words.NET.InsertBeforeOrAfter]$Table,
        [ValidateNotNullOrEmpty()]$DataTable,
        [TableDesign] $Design = [TableDesign]::ColorfulList,
        [int] $MaximumColumns = 5,
        [string[]]$Columns = @('Name', 'Value'),
        [bool] $Supress = $true
    )

    if ($DataTable.GetType().BaseType.Name -eq 'Array' -and $DataTable.GetType().Name -eq 'Object[]') {
        Write-Verbose 'Add-WordTable - Converting Array of Objects'
        $DataTable = $DataTable.ForEach( {[PSCustomObject]$_})
    }
    $ObjectType = $DataTable.GetType().Name
    Write-Verbose "Add-WordTable - Table row count: $(Get-ObjectCount $DataTable)"
    Write-Verbose "Add-WordTable - Object Type: $ObjectType"
    Write-Verbose "Add-WordTable - BaseType.Name: $($DataTable.GetType().BaseType.Name)"
    Write-Verbose "Add-WordTable - GetType Before Conversion: $ObjectType"

    If ($ObjectType -eq 'Hashtable' -or $ObjectType -eq 'OrderedDictionary') {

    } else {
        $DataTable = $DataTable | Select-Object *
    }

    $ObjectType = $DataTable.GetType().Name

    Write-Verbose "Add-WordTable - GetType After Conversion: $ObjectType"

    if ($ObjectType -eq 'Hashtable' -or $ObjectType -eq 'OrderedDictionary') {
        Write-Verbose 'Add-WordTable - Option 1'
        $NumberRows = $DataTable.Count + 1
        $NumberColumns = 2

        Write-Verbose "Add-WordTable - Column Count $($NumberColumns) Rows Count $NumberRows "
        Write-Verbose "Add-WordTable - Titles: $([string] $Columns)"

        $WordTable = New-WordTable -WordDocument $WordDocument -Paragraph $Paragraph -NrRows $NumberRows -NrColumns $NumberColumns -Supress $false

        Add-WordTableTitle -Title $Columns -Table $WordTable -MaximumColumns $MaximumColumns
        $Row = 1
        foreach ($TableEntry in $DataTable.GetEnumerator()) {
            $ColumnNrForTitle = 0
            $ColumnNrForData = 1
            $Data = Add-WordTableCellValue -Table $WordTable -Row $Row -Column $ColumnNrForTitle -Value $TableEntry.Name
            $Data = Add-WordTableCellValue -Table $WordTable -Row $Row -Column $ColumnNrForData -Value $TableEntry.Value
            Write-Verbose "Add-WordTable - RowNr: $Row / ColumnNr: $ColumnTitle Name: $($TableEntry.Name) Value: $($TableEntry.Value)"
            $Row++

        }
    } elseif ($ObjectType -eq 'PSCustomObject') {
        Write-Verbose 'Add-WordTable - Option 2'

        $Titles = Get-ObjectTitles -Object $DataTable[0]

        $NumberRows = $Titles.Count + 1
        $NumberColumns = 2

        Write-Verbose "Add-WordTable - Column Count $($NumberColumns) Rows Count $NumberRows "
        Write-Verbose "Add-WordTable - Titles: $([string] $Titles)"

        $WordTable = New-WordTable -WordDocument $WordDocument -Paragraph $Paragraph -NrRows $NumberRows -NrColumns $NumberColumns -Supress $false

        Add-WordTableTitle -Title $Columns -Table $WordTable -MaximumColumns $MaximumColumns
        $Row = 1
        foreach ($Title in $Titles) {
            $Value = Get-ObjectData -Object $DataTable -Title $Title -DoNotAddTitles

            $ColumnTitle = 0
            $ColumnData = 1
            $Data = Add-WordTableCellValue -Table $WordTable -Row $Row -Column $ColumnTitle -Value $Title
            $Data = Add-WordTableCellValue -Table $WordTable -Row $Row -Column $ColumnData -Value $Value
            Write-Verbose "Add-WordTable - Title:  $Title Value: $Value Row: $Row "
            $Row++

        }
    } elseif ($DataTable.GetType().Name -eq 'Object[]') {
        write-verbose 'Add-WordTable - option 3'

        $Titles = Get-ObjectTitles -Object $DataTable[0]

        $NumberColumns = if ($Titles.Count -ge $MaximumColumns) { $MaximumColumns } else { $Titles.Count }
        $NumberRows = $DataTable.Count + 1

        Write-Verbose "Add-WordTable - Column Count $($NumberColumns) Rows Count $NumberRows "
        Write-Verbose "Add-WordTable - Titles: $([string] $Titles)"
        #Write-Color "Column Count ", $NumberColumns, " Rows Count ", $NumberRows -C Yellow, Green, Yellow, Green

        $WordTable = New-WordTable -WordDocument $WordDocument -Paragraph $Paragraph -NrRows $NumberRows -NrColumns $NumberColumns -Supress $false

        Add-WordTableTitle -Title $Titles -Table $WordTable -MaximumColumns $MaximumColumns

        for ($b = 0; $b -lt $NumberRows - 1; $b++) {
            $a = 0
            foreach ($Title in $Titles) {
                $Data = Add-WordTableCellValue -Table $WordTable -Row $($b + 1) -Column $a -Value $DataTable[$b].$Title
                if ($a -eq $($MaximumColumns - 1)) { break; } # prevents display of more columns then there is space, choose carefully
                $a++
            }
        }
    } else {
        Write-Verbose 'Add-WordTable - Option 4'
        $pattern = 'string|bool|byte|char|decimal|double|float|int|long|sbyte|short|uint|ulong|ushort'
        $Columns = ($DataTable | Get-Member | Where-Object { $_.MemberType -like "*Property" -and $_.Definition -match $pattern }) | Select-Object Name
        #$Columns
        $NumberColumns = if ($Columns.Count -ge $MaximumColumns) { $MaximumColumns } else { $Columns.Count }
        $NumberRows = $DataTable.Count

        Write-Verbose "Add-WordTable - Column Count $($NumberColumns) Rows Count $NumberRows "
        #Write-Color "Column Count ", $NumberColumns, " Rows Count ", $NumberRows -C Yellow, Green, Yellow, Green

        $WordTable = New-WordTable -WordDocument $WordDocument -Paragraph $Paragraph -NrRows $NumberRows -NrColumns $NumberColumns -Supress $false

        $Titles = Add-WordTableTitle -Title $Columns -Table $WordTable -MaximumColumns $MaximumColumns

        for ($b = 1; $b -lt $NumberRows; $b++) {
            $a = 0
            foreach ($Title in $Columns.Name) {
                $Data = Add-WordTableCellValue -Table $WordTable -Row $b -Column $a -Value $DataTable[$b].$Title
                if ($a -eq $($MaximumColumns - 1)) { break; } # prevents display of more columns then there is space, choose carefully
                $a++


            }
        }

    }

    $WordTable.Design = $Design


    if ($Supress -eq $false) { return $WordTable } else { return }
}

function Remove-WordTable {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.InsertBeforeOrAfter] $Table
    )
    if ($Table -ne $null) {
        $Table.Remove()
    }

}

function New-WordTable {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
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

function Get-WordTable {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
        [switch] $ListTables,
        [switch] $LastTable,
        [nullable[int]] $TableID
    )
    if ($LastTable) {
        $Tables = $WordDocument.Tables
        $Table = $Tables[$Tables.Count - 1]
        return $Table
    }
    if ($ListTables) {
        return  $WordDocument.Tables
    }
    if ($TableID -ne $null) {
        return $WordDocument.Tables[$TableID]
    }
}

function Copy-WordTable {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        $TableFrom
    )
}

<#
public Table AddTable( int rowCount, int columnCount )
public new Table InsertTable( int rowCount, int columnCount )
public new Table InsertTable( int index, Table t )

#>