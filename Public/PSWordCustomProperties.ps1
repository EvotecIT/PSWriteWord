function Add-WordCustomProperty {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container]$WordDocument,
        [string] $Name,
        [string] $Value,
        [bool] $Supress
    )
    $CustomProperty = New-Object -TypeName Xceed.Words.NET.CustomProperty -ArgumentList $Name, $Value
    $Data = $WordDocument.AddCustomProperty($CustomProperty)
    if ($Supress) { return } else { return $Data }
}

function Get-WordCustomProperty {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container]$WordDocument,
        [string] $Name
    )
    if ($Property -eq $null) {
        $Data = $WordDocument.CustomProperties.Values
    } else {
        $Data = $WordDocument.CustomProperties.$Name.Value
    }
    return $Data
}