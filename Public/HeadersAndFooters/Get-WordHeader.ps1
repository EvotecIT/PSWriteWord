function Get-WordHeader {
    [CmdletBinding()]
    param (
        [Container]$WordDocument,
        [ValidateSet('All', 'First', 'Even', 'Odd')][string] $Type = 'All',
        [bool] $Supress = $false
    )
    if ($Type -eq 'All') {
        $WordDocument.Headers
    } else {
        $WordDocument.Headers.$Type
    }
}