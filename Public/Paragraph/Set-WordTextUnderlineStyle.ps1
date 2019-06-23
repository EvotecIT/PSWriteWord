function Set-WordTextUnderlineStyle {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [nullable[UnderlineStyle]] $UnderlineStyle,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $UnderlineStyle) {
        $Paragraph = $Paragraph.UnderlineStyle($UnderlineStyle)
    }
    if ($Supress) { return } else { return $Paragraph }
}