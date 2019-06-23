function Set-WordTextIndentationFirstLine {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [nullable[single]] $IndentationFirstLine,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $IndentationFirstLine -ne $null) {
        $Paragraph.IndentationFirstLine = $IndentationFirstLine
    }
    if ($Supress) { return } else { return $Paragraph }
}