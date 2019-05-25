function Add-WordCustomProperty {
    [CmdletBinding()]
    param (
        [Container]$WordDocument,
        [string] $Name,
        [string] $Value,
        [bool] $Supress
    )
    $CustomProperty = New-Object -TypeName CustomProperty -ArgumentList $Name, $Value
    $Data = $WordDocument.AddCustomProperty($CustomProperty)
    if ($Supress) { return } else { return $Data }
}