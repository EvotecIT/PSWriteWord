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
    $ObjectType = $DataTable.GetType().Name
    Write-Verbose "Convert-ObjectToProcess - GetType Before Conversion: $ObjectType"
    #$($DataTable.GetType().BaseType.Name)
    #$($DataTable.GetType().Name)
    if ($($DataTable.GetType().BaseType.Name) -eq 'Array' -and $($DataTable.GetType().Name) -eq 'Object[]') {
        Write-Verbose 'Convert-ObjectToProcess - Converting Array of Objects'
        #if ($DataTable.Count -gt 1) {
        $DataTable = $DataTable.ForEach( {[PSCustomObject]$_})
        #}

    }

    $ObjectType = $DataTable.GetType().Name
    Write-Verbose "Convert-ObjectToProcess - Table row count: $(Get-ObjectCount $DataTable)"
    Write-Verbose "Convert-ObjectToProcess - Object Type: $ObjectType"
    Write-Verbose "Convert-ObjectToProcess - BaseType.Name: $($DataTable.GetType().BaseType.Name)"
    Write-Verbose "Convert-ObjectToProcess - GetType Before Final Conversion: $ObjectType"
    If ($ObjectType -eq 'Hashtable' -or $ObjectType -eq 'OrderedDictionary' -or $ObjectType -eq 'PSCustomObject') {
        Write-Verbose 'Convert-ObjectToProcess - Skipping select for Hashtable / OrderedDictionary / PSCustomObject'
    } else {
        #if ($ObjectType -eq 'PSCustomObject') {
        #    Write-Verbose 'Convert-ObjectToProcess - Skipping all objects'
        #$DataTable = [rray] ($DataTable | Select-Object *)
        #} else {


        if ($ObjectType -eq 'Collection`1' -and $(Get-ObjectCount $DataTable) -eq 1) {
            Write-Verbose 'Convert-ObjectToProcess - Selecting all objects, returning array'
            $DataTable = [array] ($DataTable | Select-Object *)
        } else {
            Write-Verbose 'Convert-ObjectToProcess - Selecting all objects'
            $DataTable = ($DataTable | Select-Object *)
        }
        #}
    }

    $ObjectType = $DataTable.GetType().Name

    Write-Verbose "Convert-ObjectToProcess - GetType After Conversion: $ObjectType"
    return , $DataTable
}