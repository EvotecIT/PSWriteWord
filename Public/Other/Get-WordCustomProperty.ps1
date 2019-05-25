function Get-WordCustomProperty {
    [CmdletBinding()]
    param (
        [Container]$WordDocument,
        [string] $Name
    )
    if ($Property -eq $null) {
        $Data = $WordDocument.CustomProperties.Values
    } else {
        $Data = $WordDocument.CustomProperties.$Name.Value
    }
    return $Data
}