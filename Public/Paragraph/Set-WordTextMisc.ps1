function Set-WordTextMisc {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[Xceed.Document.NET.Misc]] $Misc,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $Misc) {
        $Paragraph = $Paragraph.Misc($Misc)
    }
    if ($Supress) { return } else { return $Paragraph }
}