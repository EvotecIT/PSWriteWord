function Add-WordCustomProperty {
    [CmdletBinding()]
    param (
        [Xceed.Document.NET.Container]$WordDocument,
        [string] $Name,
        [string] $Value,
        [bool] $Supress
    )
    $CustomProperty = [Xceed.Document.NET.CustomProperty]::new($Name,$Value)
    $Data = $WordDocument.AddCustomProperty($CustomProperty)
    if ($Supress) { return } else { return $Data }
}