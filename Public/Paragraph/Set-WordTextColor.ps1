function Set-WordTextColor {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [alias ("C")] [nullable[System.Drawing.Color]] $Color,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $Color -ne $null) {
        $Paragraph = $Paragraph.Color($Color)
    }
    if ($Supress) { return } else { return $Paragraph }
}