function Remove-WordPicture {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [int] $PictureID,
        [bool] $Supress
    )
    if ($null -ne $Paragraph.Pictures[$PictureID]) {
        $Paragraph.Pictures[$PictureID].Remove()
    }
    if ($supress) { return } else { return $Paragraph}
}