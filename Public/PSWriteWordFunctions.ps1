
function New-WordDocument ($FilePath = "") {
    $Word = [Xceed.Words.NET.DocX]
    $WordDocument = $Word::Create($FilePath)
    return $WordDocument
}

function Save-WordDocument ($WordDocument, $FilePath = "") {
    if ($FilePath -eq "") {
        $WordDocument.Save()
    } else {
        $WordDocument.SaveAs($FilePath)
    }
    # return $WordDocument
}

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
        $Design = "ColorfulList",
        $MaximumColumns = 5
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
function Add-List {
    param (
        $WordDocument,
        [ValidateSet('Numbered', 'Bulleted')] $ListType,
        [string[]] $ListData = $null,
        $Object = $null
    )
    $LevelPrimary = 0
    $LevelSecondary = 1
    $LevelThird = 2
    if ($ListData -ne $null) {
        $ListCount = ($ListData | Measure-Object).Count
        If ($ListCount -gt 0) {
            $List = $WordDocument.AddList($ListData[0], 0, $ListType)
            for ($i = 1; $i -lt $ListData.Count; $i++ ) {
                $WordDocument.AddListItem($List, $ListData[$i]) | Out-Null
            }
            $WordDocument.InsertList($List) | Out-Null
        }
    }
    if ($Object -ne $null) {

        $IsFirstTitle = $True
        $Titles = Get-ObjectTitles -Object $Object
        foreach ($Title in $Titles) {
            $Values = Get-ObjectData -Object $Object -Title $Title
            #$Values
            $IsFirstValue = $True
            foreach ($Value in $Values) {
                if ($IsFirstTitle -eq $True) {
                    $List = $WordDocument.AddList($Value, $LevelPrimary, $ListType)
                } else {
                    #Write-Color 'Value IsFirstTitle ', $IsFirstTitle, ' Value IsFirstValue ', $IsFirstValue, ' Count ', $Values.Count, ' Value: ', $Value -Color Yellow, Green, Yellow, Green, White, Yellow
                    if ($IsFirstValue -eq $True) {
                        $WordDocument.AddListItem($List, $Value, $LevelPrimary) | Out-Null
                    } else {
                        $WordDocument.AddListItem($List, $Value, $LevelSecondary) | Out-Null
                    }
                }
                $IsFirstTitle = $false
                $IsFirstValue = $false
            }
        }
        $WordDocument.InsertList($List) | Out-Null
    }


    <#
        foreach ($item in $HashData.GetEnumerator()) {
            #$item.Key
            #$item.value
            $entry = "$($item.Key) - $($item.Value)"
            if ($count -eq 0) {
                $List = $WordDocument.AddList($entry, 0, $ListType)
            } else {
                $WordDocument.AddListItem($List, $entry) | Out-Null
            }

            $count++
        }
          $WordDocument.InsertList($List) | Out-Null
          #>
}