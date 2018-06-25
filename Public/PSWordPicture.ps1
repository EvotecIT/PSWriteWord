function Add-WordPicture {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [Xceed.Words.NET.DocXElement] $Picture,
        [string] $ImagePath,
        [int] $Rotation,
        [switch] $FlipHorizontal,
        [switch] $FlipVertical,
        [int] $ImageWidth,
        [int] $ImageHeight,
        [string] $Description,
        [bool] $Supress = $false
    )
    if ([string]::IsNullOrEmpty($Paragraph)) {
        $Paragraph = Add-WordParagraph -WordDocument $WordDocument -Supress $false
    }
    $Image = $WordDocument.AddImage($FilePathImage )

    if ($Picture -eq $null) {
        $Picture = $Image.CreatePicture()
    }
    if ($Rotation -ne 0) { $Picture.Rotation = $Rotation }
    if ($FlipHorizontal -ne $false) { $Picture.FlipHorizontal = $FlipHorizontal }
    if ($FlipVertical -ne $false) { $Picture.FlipVertical = $FlipVertical }
    if (-not [string]::IsNullOrEmpty($Description)) { $Picture.Description = $Description }
    if ($ImageWidth -ne 0) { $Picture.Width = $ImageWidth }
    if ($ImageHeight -ne 0) { $Picture.Height = $ImageHeight }
    $data = $Paragraph.AppendPicture($Picture)

    if ($Supress) { return $data } else { return }
}

function Get-WordPicture {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph,
        [switch] $ListParagraphs,
        [int] $PictureID
    )
    if ($ListParagraphs) {
        return $WordDocument.Pictures
    } else {
        return $WordDocument.Pictures[$PictureID]
    }
}