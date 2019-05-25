function Add-WordFooter {
    [CmdletBinding()]
    param (
        [Container]$WordDocument,
        [nullable[bool]] $DifferentFirstPage,
        [nullable[bool]] $DifferentOddAndEvenPages,
        [bool] $Supress = $false
    )
    $WordDocument.AddFooters()
    if ($DifferentOddAndEvenPages -ne $null ) { $WordDocument.DifferentFirstPage = $DifferentFirstPage }
    if ($DifferentOddAndEvenPages -ne $null ) { $WordDocument.DifferentOddAndEvenPages = $DifferentOddAndEvenPages }

    if ($Supress) { return } else { return $WordDocument.Footers }
}