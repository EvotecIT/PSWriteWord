function Set-WordTextStrikeThrough {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [nullable[StrikeThrough]] $StrikeThrough,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $StrikeThrough) {
        $Paragraph = $Paragraph.StrikeThrough($StrikeThrough)
    }
    if ($Supress) { return } else { return $Paragraph }
}