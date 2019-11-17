function Add-WordHyperLink {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)] [Xceed.Document.NET.Container]$WordDocument,
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)] [Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [string] $UrlText,
        [uri] $UrlLink,
        [alias ("C")] [System.Drawing.KnownColor]$Color,
        [alias ("S")] [double] $FontSize,
        [alias ("FontName")] [string] $FontFamily,
        [alias ("B")] [nullable[bool]] $Bold,
        [alias ("I")] [nullable[bool]] $Italic,
        [alias ("U")] [Xceed.Document.NET.UnderlineStyle] $UnderlineStyle,
        [alias ('UC')] [System.Drawing.KnownColor]$UnderlineColor,
        [alias ("SA")] [double] $SpacingAfter,
        [alias ("SB")] [double] $SpacingBefore,
        [alias ("SP")] [double] $Spacing,
        [alias ("H")] [Xceed.Document.NET.Highlight] $Highlight,
        [alias ("CA")] [Xceed.Document.NET.CapsStyle] $CapsStyle,
        [alias ("ST")] [Xceed.Document.NET.StrikeThrough] $StrikeThrough,
        [alias ("HT")] [Xceed.Document.NET.HeadingType] $HeadingType,
        [int] $PercentageScale, # "Value must be one of the following: 200, 150, 100, 90, 80, 66, 50 or 33"
        [Xceed.Document.NET.Misc] $Misc,
        [string] $Language,
        [int]$Kerning, # "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72"
        [nullable[bool]]$Hidden,
        [int]$Position, #  "Value must be in the range -1585 - 1585"
        [nullable[bool]]$NewLine,
        [single] $IndentationFirstLine,
        [single] $IndentationHanging,
        [Xceed.Document.NET.Alignment] $Alignment,
        [Xceed.Document.NET.Direction] $Direction,
        [Xceed.Document.NET.ShadingType] $ShadingType,
        [System.Drawing.KnownColor]$ShadingColor,
        [Xceed.Document.NET.Script] $Script,
        [bool] $Supress = $false
    )
    $HyperLink = $WordDocument.AddHyperlink( $UrlText, $UrlLink)
    if (-not $Paragraph) {
        $Paragraph = $WordDocument.InsertParagraph()
    }
    if ($Paragraph -and $HyperLink) {
        $Data = $Paragraph.AppendHyperlink($HyperLink)
    }
    if ($Color) {
        $Paragraph = $Paragraph | Set-WordTextColor -Color $Color -Supress $false
    }
    if ($FontSize) {
        $Paragraph = $Paragraph | Set-WordTextFontSize -FontSize $FontSize -Supress $false
    }
    if ($FontFamily) {
        $Paragraph = $Paragraph | Set-WordTextFontFamily -FontFamily $FontFamily -Supress $false
    }
    if ($Bold) {
        $Paragraph = $Paragraph | Set-WordTextBold -Bold $Bold -Supress $false
    }
    if ($Italic) {
        $Paragraph = $Paragraph | Set-WordTextItalic -Italic $Italic -Supress $false
    }
    if ($UnderlineColor) {
        $Paragraph = $Paragraph | Set-WordTextUnderlineColor -UnderlineColor $UnderlineColor -Supress $false
    }
    if ($UnderlineStyle) {
        $Paragraph = $Paragraph | Set-WordTextUnderlineStyle -UnderlineStyle $UnderlineStyle -Supress $false
    }
    if ($SpacingAfter) {
        $Paragraph = $Paragraph | Set-WordTextSpacingAfter -SpacingAfter $SpacingAfter -Supress $false
    }
    if ($SpacingBefore) {
        $Paragraph = $Paragraph | Set-WordTextSpacingBefore -SpacingBefore $SpacingBefore -Supress $false
    }
    if ($Spacing) {
        $Paragraph = $Paragraph | Set-WordTextSpacing -Spacing $Spacing -Supress $false
    }
    if ($Highlight) {
        $Paragraph = $Paragraph | Set-WordTextHighlight -Highlight $Highlight -Supress $false
    }
    if ($CapsStyle) {
        $Paragraph = $Paragraph | Set-WordTextCapsStyle -CapsStyle $CapsStyle -Supress $false
    }
    if ($StrikeThrough) {
        $Paragraph = $Paragraph | Set-WordTextStrikeThrough -StrikeThrough $StrikeThrough -Supress $false
    }
    if ($PercentageScale) {
        $Paragraph = $Paragraph | Set-WordTextPercentageScale -PercentageScale $PercentageScale -Supress $false
    }
    if ($Language) {
        $Paragraph = $Paragraph | Set-WordTextLanguage -Language $Language -Supress $false
    }
    if ($Kerning) {
        $Paragraph = $Paragraph | Set-WordTextKerning -Kerning $Kerning -Supress $false
    }
    if ($Misc) {
        $Paragraph = $Paragraph | Set-WordTextMisc -Misc $Misc -Supress $false
    }
    if ($Position) {
        $Paragraph = $Paragraph | Set-WordTextPosition -Position $Position -Supress $false
    }
    if ($Hidden) {
        $Paragraph = $Paragraph | Set-WordTextHidden -Hidden $Hidden -Supress $false
    }
    if ($ShadingColor) {
        $Paragraph = $Paragraph | Set-WordTextShadingType -ShadingColor $ShadingColor -ShadingType $ShadingType -Supress $false
    }
    if ($Script) {
        $Paragraph = $Paragraph | Set-WordTextScript -Script $Script -Supress $false
    }
    if ($HeadingType) {
        $Paragraph = $Paragraph | Set-WordTextHeadingType -HeadingType $HeadingType -Supress $false
    }
    if ($IndentationFirstLine) {
        $Paragraph = $Paragraph | Set-WordTextIndentationFirstLine -IndentationFirstLine $IndentationFirstLine -Supress $false
    }
    if ($IndentationHanging) {
        $Paragraph = $Paragraph | Set-WordTextIndentationHanging -IndentationHanging $IndentationHanging -Supress $false
    }
    if ($Alignment) {
        $Paragraph = $Paragraph | Set-WordTextAlignment -Alignment $Alignment -Supress $false
    }
    if ($Direction) {
        $Paragraph = $Paragraph | Set-WordTextDirection -Direction $Direction -Supress $false
    }
    if ($Supress -eq $false) { return $Data } else { return }
}