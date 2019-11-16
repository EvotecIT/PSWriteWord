function Remove-WordParagraph {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [bool] $TrackChanges
    )
    $Paragraph.Remove($TrackChanges)
}