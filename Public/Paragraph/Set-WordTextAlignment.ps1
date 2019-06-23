function Set-WordTextAlignment {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [nullable[Alignment]] $Alignment,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $Alignment) {
        $Paragraph.Alignment = $Alignment
    }
    if ($Supress) { return } else { return $Paragraph }
}