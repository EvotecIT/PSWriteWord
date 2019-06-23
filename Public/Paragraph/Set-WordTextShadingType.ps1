function Set-WordTextShadingType {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [nullable[ShadingType]] $ShadingType,
        [nullable[System.Drawing.Color]] $ShadingColor,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $ShadingType -and $ShadingColor -ne $null) {
        $Paragraph = $Paragraph.Shading($ShadingColor, $ShadingType)
    }
    if ($Supress) { return } else { return $Paragraph }
}