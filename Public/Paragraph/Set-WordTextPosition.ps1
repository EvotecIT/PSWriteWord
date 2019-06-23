function Set-WordTextPosition {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [nullable[int]]$Position,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $Position -ne $null) {
        $Paragraph = $Paragraph.Position($Position)
    }
    if ($Supress) { return } else { return $Paragraph }
}