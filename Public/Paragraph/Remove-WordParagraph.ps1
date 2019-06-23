function Remove-WordParagraph {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [bool] $TrackChanges
    )
    $Paragraph.Remove($TrackChanges)
}