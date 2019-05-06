function Set-WordTableRowMergeCells {
    [CmdletBinding()]
    param(
        [Xceed.Words.NET.InsertBeforeOrAfter] $Table,
        [nullable[int]] $RowNr,
        [nullable[int]] $ColumnNrStart,
        [nullable[int]] $ColumnNrEnd,
        [switch] $MergeAll,
        [switch] $TrackChanges,
        [switch] $TextMerge, # Merges Text, otherwise leaves only text from 1st column
        [string] $Separator = ' ',
        [bool] $Supress = $false
    )
    if ($Table) {
        if ($MergeAll -and $RowNr -ne $null) {
            $CellsCount = $Table.Rows[$RowNr].Cells.Count
            $Table.Rows[$RowNr].MergeCells(0, $CellsCount)
            for ($paragraph = 1; $paragraph -le $Table.Rows[$RowNr].Paragraphs.Count; $paragraph++) {
                $Table.Rows[$RowNr].Paragraphs[$paragraph].Remove($TrackChanges)
            }
        } elseif ($RowNr -ne $null -and $ColumnNrStart -ne $null -and $ColumnNrEnd -ne $null) {
            $CurrentParagraphCount = $Table.Rows[$RowNr].Cells[$ColumnNrStart].Paragraphs.Count
            $Table.Rows[$RowNr].MergeCells($ColumnNrStart, $ColumnNrEnd)
            if ($TextMerge) {
                [string] $Texts = foreach ($Paragraph in $Table.Rows[$RowNr].Cells[$ColumnNrStart].Paragraphs | Select-Object -Skip ($CurrentParagraphCount - 1)) {
                    $Paragraph.Text
                } -join $Separator
            }
            # Removes Paragraphs from merged columns - Leaves only 1st column Text
            foreach ($Paragraph in $Table.Rows[$RowNr].Cells[$ColumnNrStart].Paragraphs | Select-Object -Skip $CurrentParagraphCount) {
                $Paragraph.Remove($TrackChanges)
            }
            if ($TextMerge) {
                Set-WordTextText -Paragraph $Table.Rows[$RowNr].Cells[$ColumnNrStart].Paragraphs[$CurrentParagraphCount - 1] -Text $Texts -Supress $True
            }
        }
    }
    if ($Supress) { return } else { return $Table }
}