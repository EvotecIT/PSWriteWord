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

function Convert-ObjectToProcess {
    [CmdletBinding()]
    param (
        $DataTable
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
    return $DataTable
}