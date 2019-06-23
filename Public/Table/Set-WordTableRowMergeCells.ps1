function Set-WordTableRowMergeCells {
    [CmdletBinding()]
    param(
        [InsertBeforeOrAfter] $Table,
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
        if ($MergeAll -and $null -ne $RowNr) {
            $CellsCount = $Table.Rows[$RowNr].Cells.Count
            $Table.Rows[$RowNr].MergeCells(0, $CellsCount)
            for ($paragraph = 0; $paragraph -lt $Table.Rows[$RowNr].Paragraphs.Count; $paragraph++) {
                try {
                    $Table.Rows[$RowNr].Paragraphs[$paragraph].Remove($TrackChanges)
                } catch {
                    Write-Warning -Message "Set-WordTableRowMergeCells - Failed to remove - Paragraph ($paragraph), Row ($RowNr), TrackChanges ($TrackChanges)"
                }
            }
        } elseif ($null -ne $RowNr -and $null -ne $ColumnNrStart -and $null -ne $ColumnNrEnd) {
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