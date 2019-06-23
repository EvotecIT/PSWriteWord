function Set-WordTextHighlight {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [nullable[Highlight]] $Highlight,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $Highlight) {
        $Paragraph = $Paragraph.Highlight($Highlight)
    }
    if ($Supress) { return } else { return $Paragraph }
}