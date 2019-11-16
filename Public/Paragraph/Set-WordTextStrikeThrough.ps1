function Set-WordTextStrikeThrough {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[Xceed.Document.NET.StrikeThrough]] $StrikeThrough,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $StrikeThrough) {
        $Paragraph = $Paragraph.StrikeThrough($StrikeThrough)
    }
    if ($Supress) { return } else { return $Paragraph }
}