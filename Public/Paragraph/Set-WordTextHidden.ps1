function Set-WordTextHidden {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[bool]] $Hidden,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $Hidden -ne $null) {
        $Paragraph = $Paragraph.Hidden($Hidden)
    }
    if ($Supress) { return } else { return $Paragraph }
}