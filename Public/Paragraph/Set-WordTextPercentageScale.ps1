function Set-WordTextPercentageScale {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[int]]$PercentageScale,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $PercentageScale -ne $null) {
        $Paragraph = $Paragraph.PercentageScale($PercentageScale)
    }
    if ($Supress) { return } else { return $Paragraph }
}