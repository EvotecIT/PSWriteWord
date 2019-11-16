function Set-WordTextSpacingBefore {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.InsertBeforeOrAfter] $Paragraph,
        [nullable[double]] $SpacingBefore,
        [bool] $Supress = $false
    )
    if ($null -ne $Paragraph -and $SpacingBefore -ne $null) {
        $Paragraph = $Paragraph.SpacingBefore($SpacingBefore)
    }
    if ($Supress) { return } else { return $Paragraph }
}