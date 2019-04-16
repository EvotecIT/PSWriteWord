function Remove-WordText {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [int] $Index = 0,
        [int] $Count = $($Paragraph.Text.Length),
        [bool] $TrackChanges,
        [bool] $RemoveEmptyParagraph,
        [bool] $Supress = $false
    )
    if ($Paragraph -ne $null) {
        Write-Verbose "Remove-WordText - Current text $($Paragraph.Text) "
        Write-Verbose "Remove-WordText - Removing from $Index to $Count - Paragraph Text Count: $($Paragraph.Text.Length)"
        if ($Count -ne 0) {
            $Paragraph.RemoveText($Index, $Count, $TrackChanges, $RemoveEmptyParagraph)
        }
    }
    if ($Supress) { return } else { return $Paragraph }
}