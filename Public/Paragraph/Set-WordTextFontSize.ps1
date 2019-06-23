function Set-WordTextFontSize {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [alias ("S")] [nullable[double]] $FontSize,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $FontSize -ne $null) {
        $Paragraph = $Paragraph.FontSize($FontSize)
    }
    if ($Supress) { return } else { return $Paragraph }
}