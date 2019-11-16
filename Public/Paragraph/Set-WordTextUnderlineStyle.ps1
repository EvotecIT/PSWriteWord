function Set-WordTextUnderlineStyle {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[Xceed.Document.NET.UnderlineStyle]] $UnderlineStyle,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $UnderlineStyle) {
        $Paragraph = $Paragraph.UnderlineStyle($UnderlineStyle)
    }
    if ($Supress) { return } else { return $Paragraph }
}