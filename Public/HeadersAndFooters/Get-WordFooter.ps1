function Get-WordFooter {
    [CmdletBinding()]
    param (
        [Xceed.Document.NET.Container]$WordDocument,
        [ValidateSet('All', 'First', 'Even', 'Odd')][string] $Type = 'All',
        [bool] $Supress = $false
    )
    if ($Type -eq 'All') {
        $WordDocument.Footers
    } else {
        $WordDocument.Footers.$Type
    }
}