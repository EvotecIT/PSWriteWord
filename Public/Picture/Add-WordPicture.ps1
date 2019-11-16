function Add-WordPicture {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [Xceed.Document.NET.Picture] $Picture,
        [alias('FileImagePath')][string] $ImagePath,
        [Xceed.Document.NET.Alignment] $Alignment,
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
    if ($null -eq $Picture) {
        if ($ImagePath -ne '' -and (Test-Path($ImagePath))) {
            try {
                $Image = $WordDocument.AddImage($ImagePath)
                $Picture = $Image.CreatePicture()
            } catch {
                Write-Warning "Add-WordPicture - Failed adding image. Please check with different file format/type. Aborting."
                return
            }
        } else {
            Write-Warning "Add-WordPicture - Path to ImagePath ($ImagePath) was incorrect. Aborting."
            return
        }
    }
    if ($Rotation -ne 0) { $Picture.Rotation = $Rotation }
    if ($FlipHorizontal -ne $false) { $Picture.FlipHorizontal = $FlipHorizontal }
    if ($FlipVertical -ne $false) { $Picture.FlipVertical = $FlipVertical }
    if (-not [string]::IsNullOrEmpty($Description)) { $Picture.Description = $Description }
    if ($ImageWidth -ne 0) { $Picture.Width = $ImageWidth }
    if ($ImageHeight -ne 0) { $Picture.Height = $ImageHeight }
    $Data = $Paragraph.AppendPicture($Picture)
    if ($Alignment) {
        $Data = Set-WordTextAlignment -Paragraph $Data -Alignment $Alignment
    }
    if ($Supress) { return } else { return $Data }
}