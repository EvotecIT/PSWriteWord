function Set-WordTextMisc {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [nullable[Misc]] $Misc,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $Misc) {
        $Paragraph = $Paragraph.Misc($Misc)
    }
    if ($Supress) { return } else { return $Paragraph }
}