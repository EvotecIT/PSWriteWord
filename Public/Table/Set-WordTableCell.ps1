
function Set-WordTableCell {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [nullable[int]] $RowNr,
        [nullable[int]] $ColumnNr,
        [System.Drawing.Color] $FillColor,
        [System.Drawing.Color] $ShadingColor,
        [bool] $Supress = $false
    )
    $Table = Set-WordTableCellFillColor -Table $Table -RowNr $RowNr -ColumnNr $ColumnNr -FillColor $FillColor -Supress $false
    $Table = Set-WordTableCellShadingColor  -Table $Table -RowNr $RowNr -ColumnNr $ColumnNr -ShadingColor $ShadingColor -Supress $false
    if ($Supress) { return } else { return $Table }
}
