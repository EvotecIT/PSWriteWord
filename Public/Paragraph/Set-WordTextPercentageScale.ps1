function Set-WordTextPercentageScale {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [nullable[int]]$PercentageScale,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $PercentageScale -ne $null) {
        $Paragraph = $Paragraph.PercentageScale($PercentageScale)
    }
    if ($Supress) { return } else { return $Paragraph }
}