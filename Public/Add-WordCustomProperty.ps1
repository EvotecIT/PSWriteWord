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