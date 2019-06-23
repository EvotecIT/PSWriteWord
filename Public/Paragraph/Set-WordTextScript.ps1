function Set-WordTextScript {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][InsertBeforeOrAfter] $Paragraph,
        [nullable[Script]] $Script,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $null -ne $Script) {
        $Paragraph = $Paragraph.Script($Script)
    }
    if ($Supress) { return } else { return $Paragraph }
}