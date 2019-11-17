function Set-WordTextColor {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [alias ("C")] [nullable[System.Drawing.KnownColor]] $Color,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $Color -ne $null) {
        $ConvertedColor = [System.Drawing.Color]::FromKnownColor($Color)
        $Paragraph = $Paragraph.Color($ConvertedColor)
    }
    if ($Supress) { return } else { return $Paragraph }
}