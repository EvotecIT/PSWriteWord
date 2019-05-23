function New-WordTableBorder {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.BorderStyle] $BorderStyle,
        [Xceed.Words.NET.BorderSize] $BorderSize,
        [int] $BorderSpace,
        [System.Drawing.Color] $BorderColor
    )

    $Border = New-Object -TypeName Xceed.Words.NET.Border -ArgumentList $BorderStyle, $BorderSize, $BorderSpace, $BorderColor
    return $Border
}