function Set-WordTextKerning {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [nullable[int]] $Kerning,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $Kerning -ne $null) {
        $Paragraph = $Paragraph.Kerning($Kerning)
    }
    if ($Supress) { return } else { return $Paragraph }
}