<#

Pay version of Xceed 1.5 support this:

In Section, setting the page orientation individually for the different sections will now be supported.
In Section, the following properties can now be set to configure each section of the Document: Headers, Footers, DifferentFirstPage, MarginTop, MarginBottom, MarginLeft, MarginRight, MarginHeader, MarginFooter, MirrorMargins, PageWidth, PageHeight, PageBorders, PageLayout.
In Section, the SectionBreakType property will now correctly get/set the Xml and therefore contain the desired value.

Free version (currently at 1.1 of Xceed) doesn't yet. Therefore orientation, page margins etc can only be applied globally.
#>

function Set-WordPageSettings {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [nullable[single]] $MarginLeft,
        [nullable[single]] $MarginRight,
        [nullable[single]] $MarginTop,
        [nullable[single]] $MarginBottom,
        [nullable[single]] $PageWidth,
        [nullable[single]] $PageHeight,
        [alias ("PageLayout")][nullable[Orientation]] $Orientation
    )

    Set-WordMargins -WordDocument $WordDocument -MarginLeft $MarginLeft -MarginRight $MarginRight -MarginTop $MarginTop -MarginBottom $Mar
    Set-WordPageSize -WordDocument $WordDocument -PageWidth $PageWidth -PageHeight $PageHeight
    Set-WordOrientation -WordDocument $WordDocument -Orientation $Orientation
}

function Get-WordPageSettings {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument
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

function Set-WordOrientation {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [alias ("PageLayout")][nullable[Orientation]] $Orientation
    )
    if ($Orientation -ne $null) { $WordDocument.PageLayout.Orientation = $Orientation }
}
function Set-WordMargins {
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [nullable[single]] $MarginLeft,
        [nullable[single]] $MarginRight,
        [nullable[single]] $MarginTop,
        [nullable[single]] $MarginBottom
    )

    if ($MarginLeft -ne $null) { $WordDocument.MarginLeft = $MarginLeft }
    if ($MarginRight -ne $null) {$WordDocument.MarginRight = $MarginRight }
    if ($MarginTop -ne $null) {$WordDocument.MarginTop = $MarginTop }
    if ($MarginBottom -ne $null) {$WordDocument.MarginBottom = $MarginBottom }
}
function Set-WordPageSize {
    [CmdletBinding()]
    param (
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container]$WordDocument,
        [nullable[single]] $PageWidth,
        [nullable[single]] $PageHeight
    )
    if ($PageWidth -ne $null) {$WordDocument.PageWidth = $PageWidth }
    if ($PageHeight -ne $null) {$WordDocument.PageHeight = $PageHeight }
}
