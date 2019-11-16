function Set-WordPicture {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [Xceed.Document.NET.Picture] $Picture,
        [string] $ImagePath,
        [int] $Rotation,
        [switch] $FlipHorizontal,
        [switch] $FlipVertical,
        [int] $ImageWidth,
        [int] $ImageHeight,
        [string] $Description,
        [int] $PictureID,
        [bool] $Supress = $false
    )
    $Paragraph = Remove-WordPicture -WordDocument $WordDocument -Paragraph $Paragraph -PictureID $PictureID -Supress $false
    $data = Add-WordPicture -WordDocument $WordDocument -Paragraph $Paragraph `
        -Picture $Picture `
        -ImagePath $ImagePath -ImageWidth $ImageWidth -ImageHeight $ImageHeight `
        -Rotation $Rotation -FlipHorizontal:$FlipHorizontal -FlipVertical:$FlipVertical -Supress $Supress

    if ($Supress) { return } else { return $data }
}