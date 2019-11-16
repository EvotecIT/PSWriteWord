function Set-WordTextHighlight {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[Xceed.Document.NET.Highlight]] $Highlight,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $Highlight) {
        $Paragraph = $Paragraph.Highlight($Highlight)
    }
    if ($Supress) { return } else { return $Paragraph }
}