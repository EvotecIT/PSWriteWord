function Set-WordTextBold {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[bool]] $Bold,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $Bold -ne $null -and $Bold -eq $true) {
        $Paragraph = $Paragraph.Bold()
    }
    if ($Supress) { return } else { return $Paragraph }
}