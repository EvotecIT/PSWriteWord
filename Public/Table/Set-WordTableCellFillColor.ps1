function Set-WordTableCellFillColor {
    [CmdletBinding()]
    param (
        [Xceed.Document.NET.InsertBeforeOrAfter] $Table,
        [nullable[int]] $RowNr,
        [nullable[int]] $ColumnNr,
        [nullable[System.Drawing.KnownColor]] $FillColor,
        [bool] $Supress = $false
    )

    if ($Table -and $RowNr -and $ColumnNr -and $FillColor) {
        $Cell = $Table.Rows[$RowNr].Cells[$ColumnNr]
        $ConvertedColor = [System.Drawing.Color]::FromKnownColor($FillColor)
        $Cell.FillColor = $ConvertedColor
    }
    if ($Supress) { return } else { return $Table }
}