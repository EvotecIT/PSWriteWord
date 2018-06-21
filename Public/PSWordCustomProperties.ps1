function Add-WordCustomProperty {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container]$WordDocument,
        [string] $Name,
        [string] $Value
    )
    $CustomProperty = New-Object -TypeName Xceed.Words.NET.CustomProperty -ArgumentList $Name, $Value
    $WordDocument.AddCustomProperty($CustomProperty)
}

function Get-WordCustomProperty {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container]$WordDocument,
        [string] $Name
    )
    if ($Property -eq $null) {
        return $WordDocument.CustomProperties.Values
    } else {
        return $WordDocument.CustomProperties.$Name.Value
    }
}