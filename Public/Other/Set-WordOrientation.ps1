function Set-WordOrientation {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container]$WordDocument,
        [alias ("PageLayout")][nullable[Xceed.Document.NET.Orientation]] $Orientation
    )
    if ($Orientation -ne $null) { $WordDocument.PageLayout.Orientation = $Orientation }
}