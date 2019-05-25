function Get-WordTableRow {
    [CmdletBinding()]
    param (
        [InsertBeforeOrAfter] $Table,
        [int] $RowNr,
        [int] $ColumnNr,
        [int] $ParagraphNr,
        [switch] $RowsCount
    )

    if ($Table -ne $null) {
        if ($RowsCount) {
            # returns INT - Row count number
            return $Table.Rows.Count
        }
        # returns Paragraph of a Table Row
        return $Table.Rows[$RowNr].Cells[$ColumnNr].Paragraphs[$ParagraphNr]
    }
}