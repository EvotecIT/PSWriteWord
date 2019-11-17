function Set-WordTextUnderlineColor {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[System.Drawing.KnownColor]] $UnderlineColor,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $UnderlineColor -ne $null) {
        $ConvertedColor = [System.Drawing.Color]::FromKnownColor($UnderlineColor)
        $Paragraph = $Paragraph.UnderlineColor($ConvertedColor)
    }
    if ($Supress) { return } else { return $Paragraph }
}