function Add-WordEquation {
    [CmdletBinding()]
    param (
        [Xceed.Words.NET.Container]$WordDocument,
        [string] $Equation,
        [bool] $Supress = $true
    )
    $Output = $WordDocument.InsertEquation($Equation)

    if ($Supress -eq $false) { return $Output } else { return }
}