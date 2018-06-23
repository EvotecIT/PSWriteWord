function Set-WordMargins {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container]$WordDocument,
        [nullable[single]] $MarginLeft,
        [nullable[single]] $MarginRight,
        [nullable[single]] $MarginTop,
        [nullable[single]] $MarginBottom,
        [nullable[single]] $PageWidth
    )

    if ($MarginLeft -ne $null) {
        Write-Verbose "Set-WordMargins - $MarginLeft"
        $WordDocument.MarginLeft = $MarginLeft
    }
    if ($MarginRight -ne $null) {$WordDocument.MarginRight = $MarginRight }
    if ($MarginTop -ne $null) {$WordDocument.MarginTop = $MarginTop }
    if ($MarginBottom -ne $null) {$WordDocument.MarginBottom = $MarginBottom }
    if ($PageWidth -ne $null) {$WordDocument.PageWidth = $PageWidth }
}