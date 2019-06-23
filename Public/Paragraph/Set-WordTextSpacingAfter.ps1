function Set-WordTextSpacingAfter {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [nullable[double]] $SpacingAfter,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $SpacingAfter -ne $null) {
        $Paragraph = $Paragraph.SpacingAfter($SpacingAfter)
    }
    if ($Supress) { return } else { return $Paragraph }
}