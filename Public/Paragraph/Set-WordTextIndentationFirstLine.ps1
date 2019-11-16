function Set-WordTextIndentationFirstLine {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[single]] $IndentationFirstLine,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $IndentationFirstLine -ne $null) {
        $Paragraph.IndentationFirstLine = $IndentationFirstLine
    }
    if ($Supress) { return } else { return $Paragraph }
}