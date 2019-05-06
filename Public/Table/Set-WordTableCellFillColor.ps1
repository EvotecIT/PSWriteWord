function Set-WordTableCellFillColor {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [nullable[int]] $RowNr,
        [nullable[int]] $ColumnNr,
        [nullable[System.Drawing.Color]] $FillColor,
        [bool] $Supress = $false
    )

    if ($Table -ne $null -and $RowNr -ne $null -and $ColumnNr -ne $null -and $FillColor -ne $null) {
        $Cell = $Table.Rows[$RowNr].Cells[$ColumnNr]
        $Cell.FillColor = $FillColor
    }
    if ($Supress) { return } else { return $Table }
}