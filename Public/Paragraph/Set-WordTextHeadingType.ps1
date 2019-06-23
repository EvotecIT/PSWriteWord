function Set-WordTextHeadingType {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [nullable[HeadingType]] $HeadingType,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $HeadingType) {
        #$StyleName = [string] "$HeadingType"
        Write-Verbose "Set-WordTextHeadingType - Setting StyleName to $StyleName"
        $Paragraph.StyleName = $HeadingType
    }
    if ($Supress) { return } else { return $Paragraph }
}