function Set-WordTextDirection {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [nullable[Direction]] $Direction,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $Direction) {
        $Paragraph.Direction = $Direction
    }
    if ($Supress) { return } else { return $Paragraph }
}