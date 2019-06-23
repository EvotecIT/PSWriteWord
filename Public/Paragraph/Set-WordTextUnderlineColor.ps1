function Set-WordTextUnderlineColor {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [nullable[System.Drawing.Color]] $UnderlineColor,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $UnderlineColor -ne $null) {
        $Paragraph = $Paragraph.UnderlineColor($UnderlineColor)
    }
    if ($Supress) { return } else { return $Paragraph }
}