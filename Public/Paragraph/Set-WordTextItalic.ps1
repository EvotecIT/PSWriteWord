function Set-WordTextItalic {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [nullable[bool]] $Italic,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $Italic -ne $null -and $Italic -eq $true) {
        $Paragraph = $Paragraph.Italic()
    }
    if ($Supress) { return } else { return $Paragraph }
}
