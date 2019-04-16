function Add-WordEquation {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [string] $Equation,
        [bool] $Supress = $false
    )
    $Output = $WordDocument.InsertEquation($Equation)

    if ($Supress -eq $false) { return $Output } else { return }
}