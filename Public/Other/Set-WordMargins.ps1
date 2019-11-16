function Set-WordMargins {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container]$WordDocument,
        [nullable[single]] $MarginLeft,
        [nullable[single]] $MarginRight,
        [nullable[single]] $MarginTop,
        [nullable[single]] $MarginBottom
    )

    if ($MarginLeft -ne $null) { $WordDocument.MarginLeft = $MarginLeft }
    if ($MarginRight -ne $null) {$WordDocument.MarginRight = $MarginRight }
    if ($MarginTop -ne $null) {$WordDocument.MarginTop = $MarginTop }
    if ($MarginBottom -ne $null) {$WordDocument.MarginBottom = $MarginBottom }
}