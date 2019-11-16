function Add-WordCustomProperty {
    [CmdletBinding()]
    param (
        [Xceed.Document.NET.Container]$WordDocument,
        [string] $Name,
        [string] $Value,
        [bool] $Supress
    )
    $CustomProperty = New-Object -TypeName CustomProperty -ArgumentList $Name, $Value
    $Data = $WordDocument.AddCustomProperty($CustomProperty)
    if ($Supress) { return } else { return $Data }
}