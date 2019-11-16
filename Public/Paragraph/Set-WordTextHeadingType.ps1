function Set-WordTextHeadingType {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[Xceed.Document.NET.HeadingType]] $HeadingType,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $HeadingType) {
        #$StyleName = [string] "$HeadingType"
        Write-Verbose "Set-WordTextHeadingType - Setting StyleName to $StyleName"
        $Paragraph.StyleName = $HeadingType
    }
    if ($Supress) { return } else { return $Paragraph }
}