function Set-WordTextShadingType {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[Xceed.Document.NET.ShadingType]] $ShadingType,
        [nullable[System.Drawing.KnownColor]] $ShadingColor,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $ShadingType -and $ShadingColor -ne $null) {
        $ConvertedColor = [System.Drawing.Color]::FromKnownColor($ShadingColor)
        $Paragraph = $Paragraph.Shading($ConvertedColor, $ShadingType)
    }
    if ($Supress) { return } else { return $Paragraph }
}