function Set-WordTextFontFamily {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [string] $FontFamily,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $FontFamily -ne $null -and $FontFamily -ne '') {
        $Paragraph = $Paragraph.Font($FontFamily)
    }
    if ($Supress) { return } else { return $Paragraph }
}