function Set-WordOrientation {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [alias ("PageLayout")][nullable[Xceed.Words.NET.Orientation]] $Orientation
    )
    if ($Orientation -ne $null) { $WordDocument.PageLayout.Orientation = $Orientation }
}