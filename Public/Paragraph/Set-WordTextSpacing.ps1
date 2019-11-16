function Set-WordTextSpacing {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[double]] $Spacing,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $Spacing -ne $null) {
        $Paragraph = $Paragraph.Spacing($Spacing)
    }
    if ($Supress) { return } else { return $Paragraph }
}