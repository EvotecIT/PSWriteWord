function New-WordTableBorder {
    [CmdletBinding()]
    param (
        [Xceed.Document.NET.BorderStyle] $BorderStyle,
        [Xceed.Document.NET.BorderSize] $BorderSize,
        [int] $BorderSpace,
        [System.Drawing.KnownColor] $BorderColor
    )
    $ConvertedColor = [System.Drawing.Color]::FromKnownColor($BorderColor)
    $Border = [Xceed.Document.NET.Border]::new($BorderStyle, $BorderSize, $BorderSpace, $ConvertedColor)
    return $Border
}