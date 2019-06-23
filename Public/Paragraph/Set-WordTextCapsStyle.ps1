function Set-WordTextCapsStyle {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [nullable[CapsStyle]] $CapsStyle,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $CapsStyle) {
        $Paragraph = $Paragraph.CapsStyle($CapsStyle)
    }
    if ($Supress) { return } else { return $Paragraph }
}