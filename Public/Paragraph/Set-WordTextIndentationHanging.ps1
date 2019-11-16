function Set-WordTextIndentationHanging {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[single]] $IndentationHanging,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $IndentationHanging -ne $null) {
        $Paragraph.IndentationHanging = $IndentationHanging
    }
    if ($Supress) { return } else { return $Paragraph }
}