function Set-WordOrientation {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Container]$WordDocument,
        [alias ("PageLayout")][nullable[Orientation]] $Orientation
    )
    if ($Orientation -ne $null) { $WordDocument.PageLayout.Orientation = $Orientation }
}