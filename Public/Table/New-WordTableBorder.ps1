function New-WordTableBorder {
    [CmdletBinding()]
    param (
        [Xceed.Document.NET.BorderStyle] $BorderStyle,
        [Xceed.Document.NET.BorderSize] $BorderSize,
        [int] $BorderSpace,
        [System.Drawing.Color] $BorderColor
    )

    $Border = New-Object -TypeName Border -ArgumentList $BorderStyle, $BorderSize, $BorderSpace, $BorderColor
    return $Border
}