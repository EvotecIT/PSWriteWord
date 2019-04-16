function Add-WordHeader {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container]$WordDocument,
        [nullable[bool]] $DifferentFirstPage,
        [nullable[bool]] $DifferentOddAndEvenPages,
        [bool] $Supress = $false
    )
    $WordDocument.AddHeaders()
    if ($DifferentOddAndEvenPages -ne $null ) { $WordDocument.DifferentFirstPage = $DifferentFirstPage }
    if ($DifferentOddAndEvenPages -ne $null ) { $WordDocument.DifferentOddAndEvenPages = $DifferentOddAndEvenPages }
    if ($Supress) { return } else { return $WordDocument.Headers }
}