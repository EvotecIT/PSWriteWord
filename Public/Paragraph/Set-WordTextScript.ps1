function Set-WordTextScript {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[Xceed.Document.NET.Script]] $Script,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $Script) {
        $Paragraph = $Paragraph.Script($Script)
    }
    if ($Supress) { return } else { return $Paragraph }
}