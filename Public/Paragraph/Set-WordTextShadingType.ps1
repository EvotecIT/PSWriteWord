function Set-WordTextShadingType {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[Xceed.Document.NET.ShadingType]] $ShadingType,
        [nullable[System.Drawing.Color]] $ShadingColor,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $ShadingType -and $ShadingColor -ne $null) {
        $Paragraph = $Paragraph.Shading($ShadingColor, $ShadingType)
    }
    if ($Supress) { return } else { return $Paragraph }
}