function Get-WordCustomProperty {
    [CmdletBinding()]
    param (
        [Xceed.Document.NET.Container]$WordDocument,
        [string] $Name
    )
    if ($null -eq $Property) {
        $Data = $WordDocument.CustomProperties.Values
    } else {
        $Data = $WordDocument.CustomProperties.$Name.Value
    }
    return $Data
}