function Set-WordTextAlignment {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[Xceed.Document.NET.Alignment]] $Alignment,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $Alignment) {
        $Paragraph.Alignment = $Alignment
    }
    if ($Supress) { return } else { return $Paragraph }
}