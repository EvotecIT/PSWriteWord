function Get-WordPageSettings {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Container]$WordDocument
    )

    $Object = [ordered]@{
        MarginLeft   = $WordDocument.MarginLeft
        MarginRight  = $WordDocument.MarginRight
        MarginTop    = $WordDocument.MarginTop
        MarginBottom = $WordDocument.MarginBottom
        PageWidth    = $WordDocument.PageWidth
        PageHeight   = $WordDocument.PageHeight
        Orientation  = $WordDocument.PageLayout.Orientation
    }
    return $Object
}