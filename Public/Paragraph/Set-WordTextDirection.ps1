function Set-WordTextDirection {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[Xceed.Document.NET.Direction]] $Direction,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $Direction) {
        $Paragraph.Direction = $Direction
    }
    if ($Supress) { return } else { return $Paragraph }
}