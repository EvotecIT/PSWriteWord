function Set-WordTextCapsStyle {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[Xceed.Document.NET.CapsStyle]] $CapsStyle,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $CapsStyle) {
        $Paragraph = $Paragraph.CapsStyle($CapsStyle)
    }
    if ($Supress) { return } else { return $Paragraph }
}