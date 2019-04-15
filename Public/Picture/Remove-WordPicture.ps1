function Remove-WordPicture {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [int] $PictureID,
        [bool] $Supress
    )
    if ($Paragraph.Pictures[$PictureID] -ne $null) {
        $Paragraph.Pictures[$PictureID].Remove()
    }
    if ($supress) { return } else { return $Paragraph}
}